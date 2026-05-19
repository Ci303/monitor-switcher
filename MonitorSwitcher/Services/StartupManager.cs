using System;
using Microsoft.Win32;

namespace WorkMonitorSwitcher.Services
{
    internal static class StartupManager
    {
        private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string ValueName = "MonitorSwitcher";

        public static bool IsEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
                return !string.IsNullOrWhiteSpace(key?.GetValue(ValueName) as string);
            }
            catch
            {
                return false;
            }
        }

        public static void SetEnabled(bool enabled, string executablePath)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath);
                if (enabled)
                    key?.SetValue(ValueName, $"\"{executablePath}\"", RegistryValueKind.String);
                else
                    key?.DeleteValue(ValueName, throwOnMissingValue: false);
            }
            catch
            {
                // Non-fatal. The UI setting remains persisted; Windows may block registry writes in restricted environments.
            }
        }
    }
}
