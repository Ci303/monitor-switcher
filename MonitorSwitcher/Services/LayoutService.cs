using System;
using System.Diagnostics;
using System.IO;

namespace WorkMonitorSwitcher.Services
{
    /// <summary>
    /// Thin wrapper around NirSoft MultiMonitorTool.exe for:
    /// - Saving/restoring a monitor layout (config file)
    /// - Setting the primary monitor
    /// Safe to call when the tool is missing (methods just no-op/return false).
    /// </summary>
    internal sealed class LayoutService
    {
        private readonly string _toolPath;

        public LayoutService(string toolPath)
        {
            _toolPath = toolPath ?? string.Empty;
        }

        /// <summary>
        /// Saves the current layout to the given file path.
        /// Returns true if the call was issued and the file exists afterward.
        /// </summary>
        public bool SaveLayout(string layoutPath)
        {
            if (!File.Exists(_toolPath)) return false;
            if (string.IsNullOrWhiteSpace(layoutPath)) return false;

            try
            {
                Exec("/SaveConfig", layoutPath);
                // Caller may optionally sleep/re-detect after this.
                return File.Exists(layoutPath);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Loads a saved layout from the given file path.
        /// Returns true if the call was issued (and the file existed).
        /// </summary>
        public bool LoadLayout(string layoutPath)
        {
            if (!File.Exists(_toolPath)) return false;
            if (string.IsNullOrWhiteSpace(layoutPath)) return false;
            if (!File.Exists(layoutPath)) return false;

            try
            {
                Exec("/LoadConfig", layoutPath);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Sets the primary monitor. The 'target' may be a display name or \\.\DISPLAYn.
        /// </summary>
        public void SetPrimary(string target)
        {
            if (!File.Exists(_toolPath)) return;
            if (string.IsNullOrWhiteSpace(target)) return;

            try
            {
                Exec("/SetPrimary", target);
            }
            catch
            {
                // non-fatal
            }
        }

        // ----------------- private helpers -----------------

        private const int ToolTimeoutMs = 8000;

        private void Exec(string verb, string arg)
        {
            using var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = _toolPath,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            proc.StartInfo.ArgumentList.Add(verb);
            proc.StartInfo.ArgumentList.Add(arg);

            proc.Start();

            if (!proc.WaitForExit(ToolTimeoutMs))
            {
                try { proc.Kill(entireProcessTree: true); } catch { /* ignore */ }
            }
        }
    }
}
