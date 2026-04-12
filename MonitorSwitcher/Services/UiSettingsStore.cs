using System;
using System.IO;
using System.Text.Json;
using WorkMonitorSwitcher.Model;

namespace WorkMonitorSwitcher.Services
{
    /// <summary>
    /// Persists UiSettings to a JSON file in %APPDATA%\WorkMonitorSwitcher\ui-settings.json.
    /// On first run (missing file), it defaults DarkMode from the Windows app theme.
    /// </summary>
    internal sealed class UiSettingsStore
    {
        private readonly string _path;

        public UiSettingsStore(string appDataDir)
        {
            _path = Path.Combine(appDataDir, "ui-settings.json");
        }

        public UiSettings LoadOrDefault()
        {
            try
            {
                if (File.Exists(_path))
                {
                    var json = File.ReadAllText(_path);
                    var loaded = JsonSerializer.Deserialize<UiSettings>(json);
                    if (loaded != null) return loaded;
                }
            }
            catch
            {
                // ignore and fall back
            }

            // First run / corrupted file: default to system theme (light/dark)
            return new UiSettings
            {
                DarkMode = !WindowsTheme.AppsUseLightTheme()
            };
        }

        public void Save(UiSettings settings)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_path, json);
            }
            catch
            {
                // non-fatal
            }
        }

        public string SettingsPath => _path;
    }
}
