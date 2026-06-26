using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace WorkMonitorSwitcher.Services
{
    internal sealed class LayoutProfileStore
    {
        public const string DefaultProfileName = "Default";

        private readonly string _profilesDir;
        private readonly string _indexPath;
        private readonly string _legacyLayoutPath;

        public LayoutProfileStore(string appDataDir, string legacyLayoutPath)
        {
            _profilesDir = Path.Combine(appDataDir, "layouts");
            _indexPath = Path.Combine(_profilesDir, "profiles.json");
            _legacyLayoutPath = legacyLayoutPath;
        }

        public List<string> LoadProfileNames()
        {
            try
            {
                Directory.CreateDirectory(_profilesDir);
                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { DefaultProfileName };

                if (File.Exists(_indexPath))
                {
                    var loaded = JsonSerializer.Deserialize<List<string>>(File.ReadAllText(_indexPath));
                    if (loaded != null)
                    {
                        foreach (var name in loaded.Where(n => !string.IsNullOrWhiteSpace(n)))
                            names.Add(NormalizeProfileName(name));
                    }
                }

                foreach (var file in Directory.EnumerateFiles(_profilesDir, "*.cfg"))
                    names.Add(NormalizeProfileName(Path.GetFileNameWithoutExtension(file)));

                var ordered = names
                    .OrderBy(n => n.Equals(DefaultProfileName, StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                    .ThenBy(n => n, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                SaveProfileNames(ordered);
                return ordered;
            }
            catch
            {
                return new List<string> { DefaultProfileName };
            }
        }

        public string GetLayoutPath(string? profileName)
        {
            var name = NormalizeProfileName(profileName);
            if (name.Equals(DefaultProfileName, StringComparison.OrdinalIgnoreCase))
                return _legacyLayoutPath;

            Directory.CreateDirectory(_profilesDir);
            return Path.Combine(_profilesDir, $"{SanitizeFileName(name)}.cfg");
        }

        public void AddProfileName(string? profileName)
        {
            var name = NormalizeProfileName(profileName);
            var names = LoadProfileNames();
            if (!names.Any(n => n.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                names.Add(name);
                SaveProfileNames(names);
            }
        }

        public bool DeleteProfile(string? profileName)
        {
            var name = NormalizeProfileName(profileName);
            if (name.Equals(DefaultProfileName, StringComparison.OrdinalIgnoreCase))
                return false;

            var names = LoadProfileNames()
                .Where(n => !n.Equals(name, StringComparison.OrdinalIgnoreCase))
                .ToList();
            SaveProfileNames(names);

            var path = GetLayoutPath(name);
            if (File.Exists(path))
                File.Delete(path);

            return true;
        }

        public static string NormalizeProfileName(string? profileName)
        {
            if (string.IsNullOrWhiteSpace(profileName))
                return DefaultProfileName;

            var value = profileName.Trim();
            foreach (var c in Path.GetInvalidFileNameChars())
                value = value.Replace(c, '_');

            if (value.Equals(DefaultProfileName, StringComparison.OrdinalIgnoreCase))
                return DefaultProfileName;

            var reservedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "CON", "PRN", "AUX", "NUL",
                "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
                "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            };
            if (reservedNames.Contains(value))
                value = "_" + value;

            return string.IsNullOrWhiteSpace(value) ? DefaultProfileName : value;
        }

        private void SaveProfileNames(IEnumerable<string> profileNames)
        {
            Directory.CreateDirectory(_profilesDir);
            var names = profileNames
                .Select(NormalizeProfileName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n.Equals(DefaultProfileName, StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                .ThenBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();

            AtomicFileWriter.WriteAllText(
                _indexPath,
                JsonSerializer.Serialize(names, new JsonSerializerOptions { WriteIndented = true }));
        }

        private static string SanitizeFileName(string name)
            => NormalizeProfileName(name);
    }
}
