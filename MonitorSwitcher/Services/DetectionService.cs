using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WorkMonitorSwitcher.Model;

namespace WorkMonitorSwitcher.Services
{
    /// <summary>
    /// Detects connected monitors using NirSoft MultiMonitorTool exports (CSV/TXT),
    /// with a Screen.AllScreens fallback. Produces DetectedMonitor rows with a
    /// stable key designed to survive port swaps (prefers SerialNumber / InstanceId).
    /// </summary>
    internal sealed class DetectionService
    {
        private readonly string _toolPath;

        public DetectionService(string toolPath)
        {
            _toolPath = toolPath;
        }

        /// <summary>
        /// CSV -> TXT -> Screen fallback. Always returns at least Screen fallback if possible.
        /// </summary>
        public List<DetectedMonitor> Detect()
        {
            // 1) NirSoft CSV (most structured)
            var list = DetectFromCsv();
            if (list.Count > 0) return list;

            // 2) NirSoft text export
            list = DetectFromText();
            if (list.Count > 0) return list;

            // 3) Fallback: Windows Screen API
            return DetectFromScreenApi();
        }

        // ---------------- CSV ----------------

        private List<DetectedMonitor> DetectFromCsv()
        {
            var result = new List<DetectedMonitor>();
            if (!File.Exists(_toolPath)) return result;

            string tmp = Path.GetTempFileName();
            try
            {
                ExecuteMonitorTool("/scomma", tmp);
                if (!File.Exists(tmp)) return result;

                var lines = ReadAllLinesBestEffort(tmp);
                if (lines.Length < 2) return result;

                var headers = SplitCsvLine(lines[0]);

                int idxName = IndexOf(headers, "Monitor Name", "Monitor String");
                int idxDev = IndexOf(headers, "Device Name", "DeviceName", "Name");
                int idxSerial = IndexOf(headers, "Serial Number", "SerialNumber", "Monitor Serial Number");
                int idxMonId = IndexOf(headers, "Monitor ID", "MonitorID");
                int idxInst = IndexOf(headers, "Instance ID", "InstanceID", "PNP Device ID");
                int idxActive = IndexOf(headers, "Active", "Enabled");
                int idxPosX = IndexOf(headers, "Position X", "X", "X Position", "Left-Top");
                int idxDisconnected = IndexOf(headers, "Disconnected");
                int idxMonKey = IndexOf(headers, "Monitor Key", "MonitorKey");

                foreach (var line in lines.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var cols = SplitCsvLine(line);

                    var deviceName = Get(cols, idxDev);
                    var friendlyName = Get(cols, idxName);
                    if (string.IsNullOrWhiteSpace(friendlyName))
                        friendlyName = deviceName;

                    var dm = new DetectedMonitor
                    {
                        Name = friendlyName,
                        DeviceName = deviceName,
                        SerialNumber = Get(cols, idxSerial),
                        MonitorId = Get(cols, idxMonId),
                        InstanceId = Get(cols, idxInst),
                        MonitorKey = Get(cols, idxMonKey),
                        IsActive = IsTruthy(Get(cols, idxActive)),
                        PositionX = ParsePositionX(Get(cols, idxPosX)),
                        IsPresent = !IsTruthy(Get(cols, idxDisconnected))
                    };

                    dm.StableKey = BuildStableKey(dm.SerialNumber, dm.MonitorId, dm.InstanceId, dm.MonitorKey, dm.DeviceName);
                    result.Add(dm);
                }

                ReconcileWithScreenApi(result);

                return result.Where(r => !string.IsNullOrWhiteSpace(r.DeviceName)).ToList();
            }
            catch
            {
                return result;
            }
            finally
            {
                TryDelete(tmp);
            }
        }

        // ---------------- TXT ----------------

        private List<DetectedMonitor> DetectFromText()
        {
            var result = new List<DetectedMonitor>();
            if (!File.Exists(_toolPath)) return result;

            string tmp = Path.GetTempFileName();
            try
            {
                ExecuteMonitorTool("/stext", tmp);
                if (!File.Exists(tmp)) return result;

                var lines = ReadAllLinesBestEffort(tmp);
                DetectedMonitor? cur = null;

                foreach (var line in lines)
                {
                    if (line.StartsWith("==========================================="))
                    {
                        if (cur != null)
                        {
                            cur.StableKey = BuildStableKey(cur.SerialNumber, cur.MonitorId, cur.InstanceId, cur.MonitorKey, cur.DeviceName);
                            cur.IsPresent = true;
                            result.Add(cur);
                        }
                        cur = new DetectedMonitor();
                        continue;
                    }
                    if (cur == null) continue;

                    AssignFromText(line, "Name", v => cur.DeviceName = v);
                    AssignFromText(line, "Device Name", v => cur.DeviceName = v);
                    AssignFromText(line, "Monitor Name", v =>
                    {
                        if (!string.IsNullOrWhiteSpace(v)) cur.Name = v;
                    });
                    AssignFromText(line, "Monitor String", v =>
                    {
                        if (string.IsNullOrWhiteSpace(cur.Name) && !string.IsNullOrWhiteSpace(v)) cur.Name = v;
                    });
                    AssignFromText(line, "Monitor Key", v => cur.MonitorKey = v);
                    AssignFromText(line, "Monitor ID", v => cur.MonitorId = v);
                    AssignFromText(line, "Instance ID", v => cur.InstanceId = v);
                    AssignFromText(line, "Serial Number", v => cur.SerialNumber = v);
                    AssignFromText(line, "Monitor Serial Number", v => cur.SerialNumber = v);
                    AssignFromText(line, "Active", v => cur.IsActive = v.Equals("Yes", StringComparison.OrdinalIgnoreCase));
                    AssignFromText(line, "Disconnected", v => cur.IsPresent = !v.Equals("Yes", StringComparison.OrdinalIgnoreCase));
                    AssignFromText(line, "Position X", v => cur.PositionX = ParsePositionX(v));
                    AssignFromText(line, "Left-Top", v => cur.PositionX = ParsePositionX(v));
                }
                if (cur != null)
                {
                    if (string.IsNullOrWhiteSpace(cur.Name))
                        cur.Name = cur.DeviceName;
                    cur.StableKey = BuildStableKey(cur.SerialNumber, cur.MonitorId, cur.InstanceId, cur.MonitorKey, cur.DeviceName);
                    if (!cur.IsPresent)
                        cur.IsPresent = true;
                    result.Add(cur);
                }

                ReconcileWithScreenApi(result);

                return result.Where(r => !string.IsNullOrWhiteSpace(r.DeviceName)).ToList();
            }
            catch
            {
                return result;
            }
            finally
            {
                TryDelete(tmp);
            }
        }

        // -------------- Screen API fallback --------------

        private static List<DetectedMonitor> DetectFromScreenApi()
        {
            var result = new List<DetectedMonitor>();
            foreach (var s in Screen.AllScreens)
            {
                var dm = new DetectedMonitor
                {
                    Name = s.DeviceName,
                    DeviceName = s.DeviceName,   // \\.\DISPLAYn
                    MonitorKey = string.Empty,
                    MonitorId = string.Empty,
                    InstanceId = string.Empty,
                    SerialNumber = string.Empty,
                    IsActive = true,
                    IsPresent = true,
                    PositionX = s.Bounds.X
                };
                dm.StableKey = $"DEV:{NormalizeStable(s.DeviceName)}";
                result.Add(dm);
            }
            return result;
        }

        // ---------------- Helpers ----------------
        private const int ToolTimeoutMs = 8000;
        private void ExecuteMonitorTool(string verb, string filePath)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = _toolPath,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                psi.ArgumentList.Add(verb);
                psi.ArgumentList.Add(filePath);

                using var process = new Process { StartInfo = psi };
                process.Start();

                if (!process.WaitForExit(ToolTimeoutMs))
                {
                    try { process.Kill(entireProcessTree: true); } catch { /* ignore */ }
                }
            }
            catch
            {
                // swallow; caller will fall back
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
                // ignore cleanup failures
            }
        }

        private static void AssignFromText(string line, string key, Action<string> setter)
        {
            if (!line.StartsWith(key, StringComparison.OrdinalIgnoreCase)) return;
            var idx = line.IndexOf(':');
            if (idx < 0) return;
            setter(line[(idx + 1)..].Trim());
        }

        internal static string BuildStableKey(string? serial, string? monitorId, string? instanceId, string? monitorKey, string? deviceName)
        {
            if (!string.IsNullOrWhiteSpace(serial))
                return $"SN:{serial.Trim()}";
            if (!string.IsNullOrWhiteSpace(instanceId))
                return $"IID:{NormalizeStable(instanceId)}";
            if (!string.IsNullOrWhiteSpace(monitorKey))
                return $"MK:{NormalizeStable(monitorKey)}";
            if (!string.IsNullOrWhiteSpace(monitorId))
            {
                if (!string.IsNullOrWhiteSpace(deviceName))
                    return $"MID:{NormalizeStable(monitorId)}|DEV:{NormalizeStable(deviceName)}";
                return $"MID:{NormalizeStable(monitorId)}";
            }
            if (!string.IsNullOrWhiteSpace(deviceName))
                return $"DEV:{NormalizeStable(deviceName)}";
            return $"GUID:{Guid.NewGuid():N}";
        }

        private static string NormalizeStable(string s)
        {
            var t = (s ?? string.Empty).Trim();
            t = t.Replace('/', '\\');
            while (t.Contains("\\\\")) t = t.Replace("\\\\", "\\");
            return t;
        }

        private static bool IsTruthy(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            var v = s.Trim();
            return v.Equals("Yes", StringComparison.OrdinalIgnoreCase) ||
                   v.Equals("True", StringComparison.OrdinalIgnoreCase) ||
                   v.Equals("1");
        }

        private static int IndexOf(string[] headers, params string[] options)
        {
            for (int i = 0; i < headers.Length; i++)
                if (options.Any(o => headers[i].Equals(o, StringComparison.OrdinalIgnoreCase)))
                    return i;
            return -1;
        }

        private static string Get(string[] cols, int idx)
            => (idx >= 0 && idx < cols.Length) ? cols[idx].Trim() : string.Empty;

        private static int ParseInt(string? s) => int.TryParse((s ?? "").Trim(), out var v) ? v : 0;

        private static int ParsePositionX(string? s)
        {
            var raw = (s ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw)) return 0;

            // MultiMonitorTool can emit "X, Y" for Left-Top.
            var comma = raw.IndexOf(',');
            if (comma >= 0)
                raw = raw[..comma];

            return ParseInt(raw);
        }

        private static string[] ReadAllLinesBestEffort(string path)
        {
            try
            {
                var bytes = File.ReadAllBytes(path);
                var textDefault = Encoding.Default.GetString(bytes);
                if (textDefault.Contains(",") || textDefault.Contains("===="))
                    return SplitLines(textDefault);

                var textUtf8 = Encoding.UTF8.GetString(bytes);
                return SplitLines(textUtf8);
            }
            catch
            {
                return Array.Empty<string>();
            }
        }

        private static string[] SplitLines(string s)
            => (s ?? string.Empty).Replace("\r\n", "\n").Split('\n');

        private static string[] SplitCsvLine(string line)
        {
            var list = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '\"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '\"')
                    {
                        sb.Append('\"'); i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    list.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }
            list.Add(sb.ToString());
            return list.ToArray();
        }

        /// <summary>
        /// Uses Screen.AllScreens to fix PositionX and tighten DeviceName matches
        /// (NirSoft output can sometimes be stale or incomplete).
        /// </summary>
        private static void ReconcileWithScreenApi(List<DetectedMonitor> list)
        {
            var screens = Screen.AllScreens.ToList();

            foreach (var d in list)
            {
                var s = screens.FirstOrDefault(x => x.DeviceName.Equals(d.DeviceName, StringComparison.OrdinalIgnoreCase));
                if (s != null)
                {
                    d.IsActive = true;
                    d.IsPresent = true;
                    d.PositionX = s.Bounds.X;
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(d.Name))
                {
                    s = screens.FirstOrDefault(x => d.Name.Contains(x.DeviceName, StringComparison.OrdinalIgnoreCase));
                    if (s != null)
                    {
                        d.DeviceName = s.DeviceName;
                        d.IsActive = true;
                        d.IsPresent = true;
                        d.PositionX = s.Bounds.X;
                    }
                }
            }
        }
    }
}
