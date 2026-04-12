using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using WorkMonitorSwitcher.Model;

namespace WorkMonitorSwitcher.Services
{
    /// <summary>
    /// Persists the alias map (friendly names and last-known metadata) to
    /// %APPDATA%\WorkMonitorSwitcher\monitor-aliases.json.
    /// </summary>
    internal sealed class AliasStore
    {
        private readonly string _path;

        public AliasStore(string appDataDir)
        {
            _path = Path.Combine(appDataDir, "monitor-aliases.json");
        }

        public Dictionary<string, MonitorInfo> Load()
        {
            try
            {
                if (File.Exists(_path))
                {
                    var json = File.ReadAllText(_path);
                    var dict = JsonSerializer.Deserialize<Dictionary<string, MonitorInfo>>(json);
                    if (dict != null)
                        return new Dictionary<string, MonitorInfo>(dict, StringComparer.OrdinalIgnoreCase);
                }
            }
            catch
            {
                // ignore and fall back
            }

            return new Dictionary<string, MonitorInfo>(StringComparer.OrdinalIgnoreCase);
        }

        public void Save(Dictionary<string, MonitorInfo> map)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
                var json = JsonSerializer.Serialize(map, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_path, json);
            }
            catch
            {
                // non-fatal
            }
        }

        public string AliasPath => _path;
    }
}
