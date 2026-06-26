using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace WorkMonitorSwitcher.Services
{
    internal readonly record struct DisplayPosition(int X, int Y);

    internal sealed record DisplaySourceLayout(
        string DeviceName,
        int SourceModeIndex,
        int X,
        int Y,
        int Width,
        int Height);

    internal sealed class DisplayTopologyResult
    {
        public bool Success { get; init; }
        public int? ValidateCode { get; init; }
        public int? ApplyCode { get; init; }
        public string Message { get; init; } = string.Empty;
        public IReadOnlyList<string> Details { get; init; } = Array.Empty<string>();
    }

    internal sealed class DisplayTopologyService
    {
        private const int ErrorSuccess = 0;
        private const int ErrorInsufficientBuffer = 122;

        private const uint QdcOnlyActivePaths = 0x00000002;

        private const uint SdcUseSuppliedDisplayConfig = 0x00000020;
        private const uint SdcValidate = 0x00000040;
        private const uint SdcApply = 0x00000080;
        private const uint SdcAllowChanges = 0x00000400;

        private const int DisplayConfigDeviceInfoGetSourceName = 1;
        private const uint DisplayConfigModeInfoTypeSource = 1;

        public DisplayTopologyResult DisableDisplayUsingFallbackPrimary(
            string disableDeviceName,
            string fallbackDeviceName)
        {
            if (string.IsNullOrWhiteSpace(disableDeviceName))
                return Failure("No display was supplied to disable.");
            if (string.IsNullOrWhiteSpace(fallbackDeviceName))
                return Failure("No fallback primary display was supplied.");

            try
            {
                var snapshot = QueryActiveTopology();
                var disable = snapshot.Entries.FirstOrDefault(e => DeviceNameEquals(e.Name, disableDeviceName));
                var fallback = snapshot.Entries.FirstOrDefault(e => DeviceNameEquals(e.Name, fallbackDeviceName));

                if (disable == null)
                    return Failure($"Display '{disableDeviceName}' is not active in the current topology.");
                if (fallback == null)
                    return Failure($"Fallback display '{fallbackDeviceName}' is not active in the current topology.");
                if (ReferenceEquals(disable, fallback))
                    return Failure("The display being disabled cannot also be the fallback primary display.");

                var remainingEntries = snapshot.Entries
                    .Where(e => !ReferenceEquals(e, disable))
                    .ToList();
                if (remainingEntries.Count == 0)
                    return Failure("At least one display must remain active.");

                var layouts = remainingEntries.Select(e => e.ToLayout()).ToList();
                var fallbackLayout = layouts.First(l => l.SourceModeIndex == fallback.SourceModeIndex);
                var positions = CalculateCompactedPositions(layouts, fallbackLayout);
                var orderedEntries = remainingEntries
                    .OrderBy(e => ReferenceEquals(e, fallback) ? 0 : 1)
                    .ThenBy(e => positions[e.SourceModeIndex].X)
                    .ToList();

                return ValidateAndApply(
                    snapshot,
                    orderedEntries,
                    positions,
                    $"Disabled '{disable.Name}' with '{fallback.Name}' as primary.");
            }
            catch (Exception ex)
            {
                return Failure($"Unable to update display topology: {ex.Message}");
            }
        }

        public DisplayTopologyResult SetPrimaryDisplay(string primaryDeviceName)
        {
            if (string.IsNullOrWhiteSpace(primaryDeviceName))
                return Failure("No primary display was supplied.");

            try
            {
                var snapshot = QueryActiveTopology();
                var primary = snapshot.Entries.FirstOrDefault(e => DeviceNameEquals(e.Name, primaryDeviceName));
                if (primary == null)
                    return Failure($"Display '{primaryDeviceName}' is not active in the current topology.");

                var positions = snapshot.Entries.ToDictionary(
                    e => e.SourceModeIndex,
                    e => new DisplayPosition(e.X - primary.X, e.Y - primary.Y));

                var orderedEntries = snapshot.Entries
                    .OrderBy(e => ReferenceEquals(e, primary) ? 0 : 1)
                    .ThenBy(e => positions[e.SourceModeIndex].X)
                    .ToList();

                return ValidateAndApply(
                    snapshot,
                    orderedEntries,
                    positions,
                    $"Set '{primary.Name}' as primary.");
            }
            catch (Exception ex)
            {
                return Failure($"Unable to update display topology: {ex.Message}");
            }
        }

        public DisplayTopologyResult ApplyLayoutPositionsFromConfig(string layoutPath)
        {
            if (string.IsNullOrWhiteSpace(layoutPath) || !File.Exists(layoutPath))
                return Failure("Saved layout file was not found.");

            try
            {
                var savedPositions = ReadSavedLayoutPositions(layoutPath);
                if (savedPositions.Count == 0)
                    return Failure("Saved layout file does not contain monitor positions.");

                var snapshot = QueryActiveTopology();
                var missing = snapshot.Entries
                    .Where(e => !savedPositions.ContainsKey(e.Name))
                    .Select(e => e.Name)
                    .ToList();
                if (missing.Count > 0)
                    return Failure($"Saved layout is missing active display(s): {string.Join(", ", missing)}.");

                var positions = snapshot.Entries.ToDictionary(
                    e => e.SourceModeIndex,
                    e => savedPositions[e.Name]);

                var primary = snapshot.Entries.FirstOrDefault(e =>
                {
                    var p = savedPositions[e.Name];
                    return p.X == 0 && p.Y == 0;
                }) ?? snapshot.Entries.OrderBy(e => savedPositions[e.Name].X).First();

                var orderedEntries = snapshot.Entries
                    .OrderBy(e => ReferenceEquals(e, primary) ? 0 : 1)
                    .ThenBy(e => savedPositions[e.Name].X)
                    .ToList();

                return ValidateAndApply(
                    snapshot,
                    orderedEntries,
                    positions,
                    $"Applied saved layout positions from '{layoutPath}'.");
            }
            catch (Exception ex)
            {
                return Failure($"Unable to apply saved layout positions: {ex.Message}");
            }
        }

        internal static IReadOnlyDictionary<int, DisplayPosition> CalculateCompactedPositions(
            IReadOnlyCollection<DisplaySourceLayout> remainingDisplays,
            DisplaySourceLayout fallbackPrimary)
        {
            var positions = new Dictionary<int, DisplayPosition>();
            positions[fallbackPrimary.SourceModeIndex] = new DisplayPosition(0, 0);

            int leftEdge = 0;
            foreach (var display in remainingDisplays
                         .Where(d => d.SourceModeIndex != fallbackPrimary.SourceModeIndex &&
                                     d.X < fallbackPrimary.X)
                         .OrderByDescending(d => d.X))
            {
                leftEdge -= display.Width;
                positions[display.SourceModeIndex] =
                    new DisplayPosition(leftEdge, display.Y - fallbackPrimary.Y);
            }

            int rightEdge = fallbackPrimary.Width;
            foreach (var display in remainingDisplays
                         .Where(d => d.SourceModeIndex != fallbackPrimary.SourceModeIndex &&
                                     d.X >= fallbackPrimary.X)
                         .OrderBy(d => d.X))
            {
                positions[display.SourceModeIndex] =
                    new DisplayPosition(rightEdge, display.Y - fallbackPrimary.Y);
                rightEdge += display.Width;
            }

            return positions;
        }

        private static DisplayTopologyResult ValidateAndApply(
            TopologySnapshot snapshot,
            IReadOnlyList<PathEntry> orderedEntries,
            IReadOnlyDictionary<int, DisplayPosition> positions,
            string successMessage)
        {
            foreach (var kv in positions)
            {
                var source = snapshot.Modes[kv.Key].modeInfo.sourceMode;
                source.position = new POINTL { x = kv.Value.X, y = kv.Value.Y };
                snapshot.Modes[kv.Key].modeInfo.sourceMode = source;
            }

            var paths = orderedEntries.Select(e => e.Path).ToArray();
            var details = orderedEntries
                .Select(e =>
                {
                    var p = positions[e.SourceModeIndex];
                    return $"{e.Name}: {e.X},{e.Y} -> {p.X},{p.Y}";
                })
                .ToList();

            var validateCode = SetDisplayConfig(
                (uint)paths.Length,
                paths,
                snapshot.ModeCount,
                snapshot.Modes,
                SdcUseSuppliedDisplayConfig | SdcAllowChanges | SdcValidate);
            if (validateCode != ErrorSuccess)
            {
                return new DisplayTopologyResult
                {
                    Success = false,
                    ValidateCode = validateCode,
                    Message = $"Display topology validation failed ({validateCode}).",
                    Details = details
                };
            }

            var applyCode = SetDisplayConfig(
                (uint)paths.Length,
                paths,
                snapshot.ModeCount,
                snapshot.Modes,
                SdcUseSuppliedDisplayConfig | SdcAllowChanges | SdcApply);

            return new DisplayTopologyResult
            {
                Success = applyCode == ErrorSuccess,
                ValidateCode = validateCode,
                ApplyCode = applyCode,
                Message = applyCode == ErrorSuccess
                    ? successMessage
                    : $"Display topology apply failed ({applyCode}).",
                Details = details
            };
        }

        private static TopologySnapshot QueryActiveTopology()
        {
            for (int attempt = 0; attempt < 3; attempt++)
            {
                var sizeCode = GetDisplayConfigBufferSizes(
                    QdcOnlyActivePaths,
                    out var pathCount,
                    out var modeCount);
                if (sizeCode != ErrorSuccess)
                    throw new InvalidOperationException($"GetDisplayConfigBufferSizes failed ({sizeCode}).");

                var paths = new DISPLAYCONFIG_PATH_INFO[pathCount];
                var modes = new DISPLAYCONFIG_MODE_INFO[modeCount];

                var queryCode = QueryDisplayConfig(
                    QdcOnlyActivePaths,
                    ref pathCount,
                    paths,
                    ref modeCount,
                    modes,
                    IntPtr.Zero);

                if (queryCode == ErrorInsufficientBuffer)
                    continue;
                if (queryCode != ErrorSuccess)
                    throw new InvalidOperationException($"QueryDisplayConfig failed ({queryCode}).");

                var entries = new List<PathEntry>();
                for (int i = 0; i < pathCount; i++)
                {
                    var sourceIndex = FindSourceModeIndex(paths[i], modes, modeCount);
                    if (sourceIndex < 0)
                        throw new InvalidOperationException($"Source mode was not found for path {i}.");

                    var source = modes[sourceIndex].modeInfo.sourceMode;
                    entries.Add(new PathEntry
                    {
                        Path = paths[i],
                        Name = GetSourceDeviceName(paths[i].sourceInfo),
                        SourceModeIndex = sourceIndex,
                        X = source.position.x,
                        Y = source.position.y,
                        Width = checked((int)source.width),
                        Height = checked((int)source.height)
                    });
                }

                return new TopologySnapshot(paths, modes, pathCount, modeCount, entries);
            }

            throw new InvalidOperationException("Display topology changed while it was being queried.");
        }

        private static int FindSourceModeIndex(
            DISPLAYCONFIG_PATH_INFO path,
            DISPLAYCONFIG_MODE_INFO[] modes,
            uint modeCount)
        {
            if (path.sourceInfo.modeInfoIdx < modeCount &&
                modes[path.sourceInfo.modeInfoIdx].infoType == DisplayConfigModeInfoTypeSource)
            {
                return checked((int)path.sourceInfo.modeInfoIdx);
            }

            for (int i = 0; i < modeCount; i++)
            {
                if (modes[i].infoType == DisplayConfigModeInfoTypeSource &&
                    SameLuid(modes[i].adapterId, path.sourceInfo.adapterId) &&
                    modes[i].id == path.sourceInfo.id)
                {
                    return i;
                }
            }

            return -1;
        }

        private static string GetSourceDeviceName(DISPLAYCONFIG_PATH_SOURCE_INFO source)
        {
            var request = new DISPLAYCONFIG_SOURCE_DEVICE_NAME
            {
                header =
                {
                    type = DisplayConfigDeviceInfoGetSourceName,
                    size = (uint)Marshal.SizeOf<DISPLAYCONFIG_SOURCE_DEVICE_NAME>(),
                    adapterId = source.adapterId,
                    id = source.id
                }
            };

            var code = DisplayConfigGetDeviceInfo(ref request);
            if (code != ErrorSuccess)
                throw new InvalidOperationException($"DisplayConfigGetDeviceInfo failed ({code}).");

            return request.viewGdiDeviceName ?? string.Empty;
        }

        private static Dictionary<string, DisplayPosition> ReadSavedLayoutPositions(string layoutPath)
        {
            var positions = new Dictionary<string, DisplayPosition>(StringComparer.OrdinalIgnoreCase);
            string? name = null;
            int? x = null;
            int? y = null;

            void Commit()
            {
                if (!string.IsNullOrWhiteSpace(name) && x.HasValue && y.HasValue)
                    positions[name] = new DisplayPosition(x.Value, y.Value);
            }

            foreach (var rawLine in File.ReadLines(layoutPath))
            {
                var line = rawLine.Trim();
                if (line.Length == 0 || line.StartsWith(";", StringComparison.Ordinal))
                    continue;

                if (line.StartsWith("[", StringComparison.Ordinal) &&
                    line.EndsWith("]", StringComparison.Ordinal))
                {
                    Commit();
                    name = null;
                    x = null;
                    y = null;
                    continue;
                }

                var equals = line.IndexOf('=');
                if (equals <= 0)
                    continue;

                var key = line[..equals].Trim();
                var value = line[(equals + 1)..].Trim();
                if (key.Equals("Name", StringComparison.OrdinalIgnoreCase))
                {
                    name = value;
                }
                else if (key.Equals("PositionX", StringComparison.OrdinalIgnoreCase) &&
                         int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedX))
                {
                    x = parsedX;
                }
                else if (key.Equals("PositionY", StringComparison.OrdinalIgnoreCase) &&
                         int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedY))
                {
                    y = parsedY;
                }
            }

            Commit();
            return positions;
        }

        private static bool DeviceNameEquals(string left, string right)
            => MonitorTargetResolver.TargetsEquivalent(left, right);

        private static bool SameLuid(LUID left, LUID right)
            => left.LowPart == right.LowPart && left.HighPart == right.HighPart;

        private static DisplayTopologyResult Failure(string message)
            => new() { Success = false, Message = message };

        private sealed record TopologySnapshot(
            DISPLAYCONFIG_PATH_INFO[] Paths,
            DISPLAYCONFIG_MODE_INFO[] Modes,
            uint PathCount,
            uint ModeCount,
            IReadOnlyList<PathEntry> Entries);

        private sealed class PathEntry
        {
            public DISPLAYCONFIG_PATH_INFO Path { get; init; }
            public string Name { get; init; } = string.Empty;
            public int SourceModeIndex { get; init; }
            public int X { get; init; }
            public int Y { get; init; }
            public int Width { get; init; }
            public int Height { get; init; }

            public DisplaySourceLayout ToLayout()
                => new(Name, SourceModeIndex, X, Y, Width, Height);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LUID
        {
            public uint LowPart;
            public int HighPart;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINTL
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPLAYCONFIG_RATIONAL
        {
            public uint Numerator;
            public uint Denominator;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPLAYCONFIG_2DREGION
        {
            public uint cx;
            public uint cy;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPLAYCONFIG_VIDEO_SIGNAL_INFO
        {
            public ulong pixelRate;
            public DISPLAYCONFIG_RATIONAL hSyncFreq;
            public DISPLAYCONFIG_RATIONAL vSyncFreq;
            public DISPLAYCONFIG_2DREGION activeSize;
            public DISPLAYCONFIG_2DREGION totalSize;
            public uint videoStandard;
            public uint scanLineOrdering;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPLAYCONFIG_TARGET_MODE
        {
            public DISPLAYCONFIG_VIDEO_SIGNAL_INFO targetVideoSignalInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPLAYCONFIG_SOURCE_MODE
        {
            public uint width;
            public uint height;
            public uint pixelFormat;
            public POINTL position;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPLAYCONFIG_DESKTOP_IMAGE_INFO
        {
            public POINTL PathSourceSize;
            public RECT DesktopImageRegion;
            public RECT DesktopImageClip;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct DISPLAYCONFIG_MODE_INFO_UNION
        {
            [FieldOffset(0)] public DISPLAYCONFIG_TARGET_MODE targetMode;
            [FieldOffset(0)] public DISPLAYCONFIG_SOURCE_MODE sourceMode;
            [FieldOffset(0)] public DISPLAYCONFIG_DESKTOP_IMAGE_INFO desktopImageInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPLAYCONFIG_MODE_INFO
        {
            public uint infoType;
            public uint id;
            public LUID adapterId;
            public DISPLAYCONFIG_MODE_INFO_UNION modeInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPLAYCONFIG_PATH_SOURCE_INFO
        {
            public LUID adapterId;
            public uint id;
            public uint modeInfoIdx;
            public uint statusFlags;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPLAYCONFIG_PATH_TARGET_INFO
        {
            public LUID adapterId;
            public uint id;
            public uint modeInfoIdx;
            public uint outputTechnology;
            public uint rotation;
            public uint scaling;
            public DISPLAYCONFIG_RATIONAL refreshRate;
            public uint scanLineOrdering;
            public int targetAvailable;
            public uint statusFlags;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPLAYCONFIG_PATH_INFO
        {
            public DISPLAYCONFIG_PATH_SOURCE_INFO sourceInfo;
            public DISPLAYCONFIG_PATH_TARGET_INFO targetInfo;
            public uint flags;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DISPLAYCONFIG_DEVICE_INFO_HEADER
        {
            public int type;
            public uint size;
            public LUID adapterId;
            public uint id;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct DISPLAYCONFIG_SOURCE_DEVICE_NAME
        {
            public DISPLAYCONFIG_DEVICE_INFO_HEADER header;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string? viewGdiDeviceName;
        }

        [DllImport("user32.dll")]
        private static extern int GetDisplayConfigBufferSizes(
            uint flags,
            out uint numPathArrayElements,
            out uint numModeInfoArrayElements);

        [DllImport("user32.dll")]
        private static extern int QueryDisplayConfig(
            uint flags,
            ref uint numPathArrayElements,
            [Out] DISPLAYCONFIG_PATH_INFO[] pathInfoArray,
            ref uint numModeInfoArrayElements,
            [Out] DISPLAYCONFIG_MODE_INFO[] modeInfoArray,
            IntPtr currentTopologyId);

        [DllImport("user32.dll")]
        private static extern int SetDisplayConfig(
            uint numPathArrayElements,
            [In] DISPLAYCONFIG_PATH_INFO[] pathInfoArray,
            uint numModeInfoArrayElements,
            [In] DISPLAYCONFIG_MODE_INFO[] modeInfoArray,
            uint flags);

        [DllImport("user32.dll")]
        private static extern int DisplayConfigGetDeviceInfo(
            ref DISPLAYCONFIG_SOURCE_DEVICE_NAME requestPacket);
    }
}
