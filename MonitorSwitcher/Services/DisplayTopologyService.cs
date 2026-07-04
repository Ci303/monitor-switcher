using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using WorkMonitorSwitcher.Model;

namespace WorkMonitorSwitcher.Services
{
    internal readonly record struct DisplayPosition(int X, int Y);

    internal readonly record struct DisplaySize(int Width, int Height);

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

        public DisplayTopologyResult ApplyLayoutPositionsFromConfig(
            string layoutPath,
            IReadOnlyCollection<DetectedMonitor>? detectedMonitors = null,
            IReadOnlyCollection<SavedLayoutIdentity>? savedIdentities = null)
        {
            if (string.IsNullOrWhiteSpace(layoutPath) || !File.Exists(layoutPath))
                return Failure("Saved layout file was not found.");

            try
            {
                var savedLayouts = ReadSavedLayoutSettings(layoutPath);
                if (savedLayouts.Count == 0)
                    return Failure("Saved layout file does not contain monitor positions.");

                var snapshot = QueryActiveTopology();
                var resolvedLayouts = ResolveSavedLayoutsForCurrentDevices(
                    savedLayouts,
                    detectedMonitors,
                    snapshot.Entries.Select(e => e.Name),
                    savedIdentities);

                var missing = snapshot.Entries
                    .Where(e => !resolvedLayouts.ContainsKey(e.Name))
                    .Select(e => e.Name)
                    .ToList();
                if (missing.Count > 0)
                    return Failure($"Saved layout is missing active display(s): {string.Join(", ", missing)}.");

                var positions = snapshot.Entries.ToDictionary(
                    e => e.SourceModeIndex,
                    e => resolvedLayouts[e.Name].Position);
                var sizes = snapshot.Entries
                    .Where(e => resolvedLayouts[e.Name].Size.HasValue)
                    .ToDictionary(
                        e => e.SourceModeIndex,
                        e => resolvedLayouts[e.Name].Size!.Value);
                var rotations = snapshot.Entries
                    .Where(e => resolvedLayouts[e.Name].Rotation.HasValue)
                    .ToDictionary(
                        e => e.SourceModeIndex,
                        e => resolvedLayouts[e.Name].Rotation!.Value);
                var layoutSourceDevices = snapshot.Entries.ToDictionary(
                    e => e.SourceModeIndex,
                    e => resolvedLayouts[e.Name].DeviceName);

                var primary = snapshot.Entries.FirstOrDefault(e =>
                {
                    var p = resolvedLayouts[e.Name].Position;
                    return p.X == 0 && p.Y == 0;
                }) ?? snapshot.Entries.OrderBy(e => resolvedLayouts[e.Name].Position.X).First();

                var orderedEntries = snapshot.Entries
                    .OrderBy(e => ReferenceEquals(e, primary) ? 0 : 1)
                    .ThenBy(e => resolvedLayouts[e.Name].Position.X)
                    .ToList();

                return ValidateAndApply(
                    snapshot,
                    orderedEntries,
                    positions,
                    $"Applied saved layout positions from '{layoutPath}'.",
                    sizes,
                    rotations,
                    layoutSourceDevices);
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
            string successMessage,
            IReadOnlyDictionary<int, DisplaySize>? sizes = null,
            IReadOnlyDictionary<int, uint>? rotations = null,
            IReadOnlyDictionary<int, string>? layoutSourceDevices = null)
        {
            foreach (var kv in positions)
            {
                var source = snapshot.Modes[kv.Key].modeInfo.sourceMode;
                source.position = new POINTL { x = kv.Value.X, y = kv.Value.Y };
                if (sizes != null && sizes.TryGetValue(kv.Key, out var size))
                {
                    source.width = checked((uint)size.Width);
                    source.height = checked((uint)size.Height);
                }
                snapshot.Modes[kv.Key].modeInfo.sourceMode = source;
            }

            var paths = orderedEntries.Select(e => e.Path).ToArray();
            if (rotations != null)
            {
                for (int i = 0; i < paths.Length; i++)
                {
                    var sourceModeIndex = orderedEntries[i].SourceModeIndex;
                    if (!rotations.TryGetValue(sourceModeIndex, out var rotation))
                        continue;

                    var target = paths[i].targetInfo;
                    target.rotation = rotation;
                    paths[i].targetInfo = target;
                }
            }

            var details = orderedEntries
                .Select(e =>
                {
                    var p = positions[e.SourceModeIndex];
                    var parts = new List<string> { $"{e.Name}: {e.X},{e.Y} -> {p.X},{p.Y}" };
                    if (layoutSourceDevices != null &&
                        layoutSourceDevices.TryGetValue(e.SourceModeIndex, out var savedDeviceName) &&
                        !DeviceNameEquals(e.Name, savedDeviceName))
                    {
                        parts.Add($"saved as {savedDeviceName}");
                    }
                    if (sizes != null && sizes.TryGetValue(e.SourceModeIndex, out var size))
                        parts.Add($"size {e.Width}x{e.Height} -> {size.Width}x{size.Height}");
                    if (rotations != null && rotations.TryGetValue(e.SourceModeIndex, out var rotation))
                        parts.Add($"rotation {e.Rotation} -> {rotation}");
                    return string.Join("; ", parts);
                })
                .ToList();

            if (IsTopologyAlreadyApplied(snapshot.Entries.Count, orderedEntries, positions, sizes, rotations))
            {
                return new DisplayTopologyResult
                {
                    Success = true,
                    Message = $"{successMessage} No display topology change was required.",
                    Details = details
                };
            }

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

        private static bool IsTopologyAlreadyApplied(
            int activePathCount,
            IReadOnlyList<PathEntry> orderedEntries,
            IReadOnlyDictionary<int, DisplayPosition> positions,
            IReadOnlyDictionary<int, DisplaySize>? sizes,
            IReadOnlyDictionary<int, uint>? rotations)
        {
            if (orderedEntries.Count != activePathCount)
                return false;

            if (orderedEntries.Count == 0)
                return true;

            foreach (var entry in orderedEntries)
            {
                if (!positions.TryGetValue(entry.SourceModeIndex, out var position))
                    return false;

                if (entry.X != position.X || entry.Y != position.Y)
                    return false;

                var intendedRotation = rotations != null &&
                                       rotations.TryGetValue(entry.SourceModeIndex, out var rotation)
                    ? rotation
                    : entry.Rotation;

                if (rotations != null &&
                    rotations.TryGetValue(entry.SourceModeIndex, out var expectedRotation) &&
                    entry.Rotation != expectedRotation)
                {
                    return false;
                }

                if (sizes != null &&
                    sizes.TryGetValue(entry.SourceModeIndex, out var size) &&
                    !DisplaySizeMatches(entry, size, intendedRotation))
                {
                    return false;
                }
            }

            var primary = orderedEntries[0];
            return primary.X == 0 && primary.Y == 0;
        }

        private static bool DisplaySizeMatches(PathEntry entry, DisplaySize size, uint rotation)
        {
            if (entry.Width == size.Width && entry.Height == size.Height)
                return true;

            return IsQuarterTurn(rotation) &&
                   entry.Width == size.Height &&
                   entry.Height == size.Width;
        }

        private static bool IsQuarterTurn(uint rotation)
            => rotation == 2 || rotation == 4;

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
                        Height = checked((int)source.height),
                        Rotation = paths[i].targetInfo.rotation
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

        internal static IReadOnlyDictionary<string, string> ResolveSavedLayoutDeviceNameMap(
            string layoutPath,
            IReadOnlyCollection<DetectedMonitor> detectedMonitors,
            IEnumerable<string> currentDeviceNames,
            IReadOnlyCollection<SavedLayoutIdentity>? savedIdentities = null)
        {
            var savedLayouts = ReadSavedLayoutSettings(layoutPath);
            return ResolveSavedLayoutsForCurrentDevices(savedLayouts, detectedMonitors, currentDeviceNames, savedIdentities)
                .ToDictionary(kv => kv.Key, kv => kv.Value.DeviceName, StringComparer.OrdinalIgnoreCase);
        }

        private static Dictionary<string, SavedDisplayLayout> ResolveSavedLayoutsForCurrentDevices(
            IReadOnlyDictionary<string, SavedDisplayLayout> savedLayouts,
            IReadOnlyCollection<DetectedMonitor>? detectedMonitors,
            IEnumerable<string> currentDeviceNames,
            IReadOnlyCollection<SavedLayoutIdentity>? savedIdentities = null)
        {
            var resolved = new Dictionary<string, SavedDisplayLayout>(StringComparer.OrdinalIgnoreCase);
            var usedSavedDeviceNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var detectedByDevice = BuildDetectedByDeviceName(detectedMonitors);
            var strongIdentitySidecar = HasStrongIdentitySidecar(savedIdentities, savedLayouts);

            foreach (var currentDeviceName in currentDeviceNames
                         .Where(n => !string.IsNullOrWhiteSpace(n))
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                SavedDisplayLayout? layout = null;
                var currentDeviceKey = MonitorTargetResolver.NormalizeDeviceNameForComparison(currentDeviceName);

                if (detectedByDevice.TryGetValue(currentDeviceKey, out var detected))
                {
                    layout = FindUniqueUnusedSavedLayout(
                        savedLayouts,
                        savedIdentities,
                        usedSavedDeviceNames,
                        detected);

                    layout ??= FindUniqueUnusedSavedLayout(
                        savedLayouts.Values,
                        usedSavedDeviceNames,
                        saved => IdentityValueEquals(saved.SerialNumber, detected.SerialNumber));

                    layout ??= FindUniqueUnusedSavedLayout(
                        savedLayouts.Values,
                        usedSavedDeviceNames,
                        saved => IdentityValueEquals(saved.MonitorId, detected.MonitorId));
                }

                if (layout == null &&
                    !strongIdentitySidecar &&
                    savedLayouts.TryGetValue(currentDeviceName, out var savedByDeviceName) &&
                    !usedSavedDeviceNames.Contains(savedByDeviceName.DeviceName))
                {
                    layout = savedByDeviceName;
                }

                if (layout == null)
                    continue;

                resolved[currentDeviceName] = layout;
                usedSavedDeviceNames.Add(layout.DeviceName);
            }

            return resolved;
        }

        private static Dictionary<string, DetectedMonitor> BuildDetectedByDeviceName(
            IReadOnlyCollection<DetectedMonitor>? detectedMonitors)
        {
            if (detectedMonitors == null || detectedMonitors.Count == 0)
                return new Dictionary<string, DetectedMonitor>(StringComparer.OrdinalIgnoreCase);

            return detectedMonitors
                .Where(d => d.IsPresent &&
                            d.IsActive &&
                            !string.IsNullOrWhiteSpace(d.DeviceName))
                .Select(d => new
                {
                    DeviceName = MonitorTargetResolver.NormalizeDeviceNameForComparison(d.DeviceName),
                    Monitor = d
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.DeviceName))
                .GroupBy(x => x.DeviceName, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() == 1)
                .ToDictionary(
                    g => g.Key,
                    g => g.First().Monitor,
                    StringComparer.OrdinalIgnoreCase);
        }

        private static SavedDisplayLayout? FindUniqueUnusedSavedLayout(
            IReadOnlyDictionary<string, SavedDisplayLayout> savedLayouts,
            IReadOnlyCollection<SavedLayoutIdentity>? savedIdentities,
            ISet<string> usedSavedDeviceNames,
            DetectedMonitor detected)
        {
            if (savedIdentities == null || savedIdentities.Count == 0)
                return null;

            var identity = FindUniqueUnusedSavedIdentity(
                savedLayouts,
                savedIdentities,
                usedSavedDeviceNames,
                saved => IdentityValueEquals(saved.StableKey, detected.StableKey));

            identity ??= FindUniqueUnusedSavedIdentity(
                savedLayouts,
                savedIdentities,
                usedSavedDeviceNames,
                saved => IdentityValueEquals(saved.SerialNumber, detected.SerialNumber));

            identity ??= FindUniqueUnusedSavedIdentity(
                savedLayouts,
                savedIdentities,
                usedSavedDeviceNames,
                saved => IdentityValueEquals(saved.InstanceId, detected.InstanceId));

            identity ??= FindUniqueUnusedSavedIdentity(
                savedLayouts,
                savedIdentities,
                usedSavedDeviceNames,
                saved => IdentityValueEquals(saved.MonitorKey, detected.MonitorKey));

            identity ??= FindUniqueUnusedSavedIdentity(
                savedLayouts,
                savedIdentities,
                usedSavedDeviceNames,
                saved => IdentityValueEquals(saved.MonitorId, detected.MonitorId));

            if (identity == null)
                return null;

            var layoutDeviceName = GetIdentityLayoutDeviceName(identity);
            return savedLayouts.TryGetValue(layoutDeviceName, out var layout)
                ? layout
                : null;
        }

        private static SavedLayoutIdentity? FindUniqueUnusedSavedIdentity(
            IReadOnlyDictionary<string, SavedDisplayLayout> savedLayouts,
            IEnumerable<SavedLayoutIdentity> savedIdentities,
            ISet<string> usedSavedDeviceNames,
            Func<SavedLayoutIdentity, bool> predicate)
        {
            var matches = savedIdentities
                .Where(saved =>
                {
                    var layoutDeviceName = GetIdentityLayoutDeviceName(saved);
                    return !string.IsNullOrWhiteSpace(layoutDeviceName) &&
                           savedLayouts.ContainsKey(layoutDeviceName) &&
                           !usedSavedDeviceNames.Contains(layoutDeviceName) &&
                           predicate(saved);
                })
                .GroupBy(GetIdentityLayoutDeviceName, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();

            return matches.Count == 1 ? matches[0] : null;
        }

        private static bool HasStrongIdentitySidecar(
            IReadOnlyCollection<SavedLayoutIdentity>? savedIdentities,
            IReadOnlyDictionary<string, SavedDisplayLayout> savedLayouts)
        {
            return savedIdentities != null &&
                   savedIdentities.Any(saved =>
                       savedLayouts.ContainsKey(GetIdentityLayoutDeviceName(saved)) &&
                       HasStrongIdentity(saved));
        }

        private static bool HasStrongIdentity(SavedLayoutIdentity identity)
        {
            if (NormalizeIdentityValue(identity.SerialNumber).Length > 0)
                return true;
            if (NormalizeIdentityValue(identity.InstanceId).Length > 0)
                return true;
            if (NormalizeIdentityValue(identity.MonitorKey).Length > 0)
                return true;
            if (NormalizeIdentityValue(identity.MonitorId).Length > 0)
                return true;

            var stableKey = NormalizeIdentityValue(identity.StableKey);
            return stableKey.Length > 0 &&
                   !stableKey.StartsWith("DEV:", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetIdentityLayoutDeviceName(SavedLayoutIdentity identity)
            => FirstNonBlank(identity.LayoutDeviceName, identity.DeviceName);

        private static SavedDisplayLayout? FindUniqueUnusedSavedLayout(
            IEnumerable<SavedDisplayLayout> savedLayouts,
            ISet<string> usedSavedDeviceNames,
            Func<SavedDisplayLayout, bool> predicate)
        {
            var matches = savedLayouts
                .Where(saved => !usedSavedDeviceNames.Contains(saved.DeviceName))
                .Where(predicate)
                .ToList();

            return matches.Count == 1 ? matches[0] : null;
        }

        private static string FirstNonBlank(params string?[] values)
            => values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v))?.Trim() ?? string.Empty;

        private static bool IdentityValueEquals(string? left, string? right)
        {
            var leftValue = NormalizeIdentityValue(left);
            var rightValue = NormalizeIdentityValue(right);

            return leftValue.Length > 0 &&
                   rightValue.Length > 0 &&
                   leftValue.Equals(rightValue, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeIdentityValue(string? value)
        {
            var normalized = (value ?? string.Empty).Trim().Replace('/', '\\');
            while (normalized.Contains("\\\\", StringComparison.Ordinal))
                normalized = normalized.Replace("\\\\", "\\");
            return normalized;
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

        private static Dictionary<string, SavedDisplayLayout> ReadSavedLayoutSettings(string layoutPath)
        {
            var layouts = new Dictionary<string, SavedDisplayLayout>(StringComparer.OrdinalIgnoreCase);
            string? name = null;
            string? serialNumber = null;
            string? monitorId = null;
            int? x = null;
            int? y = null;
            int? width = null;
            int? height = null;
            uint? rotation = null;

            void Commit()
            {
                if (!string.IsNullOrWhiteSpace(name) && x.HasValue && y.HasValue)
                {
                    var size = width.HasValue && height.HasValue
                        ? new DisplaySize(width.Value, height.Value)
                        : (DisplaySize?)null;
                    layouts[name] = new SavedDisplayLayout(
                        name,
                        serialNumber ?? string.Empty,
                        monitorId ?? string.Empty,
                        new DisplayPosition(x.Value, y.Value),
                        size,
                        rotation);
                }
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
                    serialNumber = null;
                    monitorId = null;
                    x = null;
                    y = null;
                    width = null;
                    height = null;
                    rotation = null;
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
                else if (key.Equals("SerialNumber", StringComparison.OrdinalIgnoreCase))
                {
                    serialNumber = value;
                }
                else if (key.Equals("MonitorID", StringComparison.OrdinalIgnoreCase))
                {
                    monitorId = value;
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
                else if (key.Equals("Width", StringComparison.OrdinalIgnoreCase) &&
                         int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedWidth))
                {
                    width = parsedWidth;
                }
                else if (key.Equals("Height", StringComparison.OrdinalIgnoreCase) &&
                         int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedHeight))
                {
                    height = parsedHeight;
                }
                else if (key.Equals("DisplayOrientation", StringComparison.OrdinalIgnoreCase) &&
                         int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedOrientation) &&
                         TryMapDisplayOrientationToCcdRotation(parsedOrientation, out var parsedRotation))
                {
                    rotation = parsedRotation;
                }
            }

            Commit();
            return layouts;
        }

        internal static bool TryMapDisplayOrientationToCcdRotation(int orientation, out uint rotation)
        {
            // MultiMonitorTool stores the same values as DEVMODE.dmDisplayOrientation.
            // CCD rotation uses 1-based DISPLAYCONFIG_ROTATION values.
            rotation = orientation switch
            {
                0 => 1, // identity / landscape
                1 => 2, // 90 degrees
                2 => 3, // 180 degrees
                3 => 4, // 270 degrees
                _ => 0
            };

            return rotation != 0;
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
            public uint Rotation { get; init; }

            public DisplaySourceLayout ToLayout()
                => new(Name, SourceModeIndex, X, Y, Width, Height);
        }

        private sealed record SavedDisplayLayout(
            string DeviceName,
            string SerialNumber,
            string MonitorId,
            DisplayPosition Position,
            DisplaySize? Size,
            uint? Rotation);

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
