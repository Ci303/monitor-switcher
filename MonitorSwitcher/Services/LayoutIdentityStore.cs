using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using WorkMonitorSwitcher.Model;

namespace WorkMonitorSwitcher.Services
{
    internal sealed class SavedLayoutIdentity
    {
        public string LayoutDeviceName { get; set; } = string.Empty;
        public string StableKey { get; set; } = string.Empty;
        public string DeviceName { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string MonitorKey { get; set; } = string.Empty;
        public string MonitorId { get; set; } = string.Empty;
        public string InstanceId { get; set; } = string.Empty;
        public string SerialNumber { get; set; } = string.Empty;
    }

    internal sealed class LayoutIdentityFile
    {
        public int Version { get; set; } = 1;
        public DateTimeOffset SavedAtUtc { get; set; } = DateTimeOffset.UtcNow;
        public List<SavedLayoutIdentity> Monitors { get; set; } = new();
    }

    internal static class LayoutIdentityStore
    {
        private const string SidecarSuffix = ".identity.json";

        public static string GetIdentityPath(string layoutPath)
            => $"{layoutPath}{SidecarSuffix}";

        public static IReadOnlyList<SavedLayoutIdentity> Load(string layoutPath)
        {
            if (string.IsNullOrWhiteSpace(layoutPath))
                return Array.Empty<SavedLayoutIdentity>();

            var path = GetIdentityPath(layoutPath);
            if (!File.Exists(path))
                return Array.Empty<SavedLayoutIdentity>();

            try
            {
                var file = JsonSerializer.Deserialize<LayoutIdentityFile>(File.ReadAllText(path));
                if (file?.Monitors == null)
                    return Array.Empty<SavedLayoutIdentity>();

                return file.Monitors
                    .Where(m => !string.IsNullOrWhiteSpace(m.LayoutDeviceName) ||
                                !string.IsNullOrWhiteSpace(m.DeviceName))
                    .ToList();
            }
            catch
            {
                return Array.Empty<SavedLayoutIdentity>();
            }
        }

        public static bool Save(string layoutPath, IEnumerable<DetectedMonitor> detectedMonitors)
        {
            if (string.IsNullOrWhiteSpace(layoutPath))
                return false;

            var monitors = detectedMonitors
                .Where(m => m.IsPresent &&
                            m.IsActive &&
                            !string.IsNullOrWhiteSpace(m.DeviceName))
                .Select(m => new SavedLayoutIdentity
                {
                    LayoutDeviceName = m.DeviceName,
                    StableKey = m.StableKey,
                    DeviceName = m.DeviceName,
                    Name = m.Name,
                    MonitorKey = m.MonitorKey,
                    MonitorId = m.MonitorId,
                    InstanceId = m.InstanceId,
                    SerialNumber = m.SerialNumber
                })
                .ToList();

            if (monitors.Count == 0)
                return false;

            var file = new LayoutIdentityFile
            {
                SavedAtUtc = DateTimeOffset.UtcNow,
                Monitors = monitors
            };

            var identityPath = GetIdentityPath(layoutPath);
            var directory = Path.GetDirectoryName(identityPath);
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);

            AtomicFileWriter.WriteAllText(
                identityPath,
                JsonSerializer.Serialize(file, new JsonSerializerOptions { WriteIndented = true }));

            return true;
        }
    }
}
