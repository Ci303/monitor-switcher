using System;
using System.Collections.Generic;

namespace WorkMonitorSwitcher.Model
{
    // Persistent per-user UI settings
    public class UiSettings
    {
        public bool DarkMode { get; set; }
        public int WindowX { get; set; }
        public int WindowY { get; set; }
        public int WindowWidth { get; set; }
        public int WindowHeight { get; set; }
        public bool AlwaysOnTop { get; set; }   // default false

    }

    // Saved metadata for a given monitor stable key
    public class MonitorInfo
    {
        public string Name { get; set; } = string.Empty;   // Friendly alias shown in UI

        // Last-seen identifiers (kept for backward compatibility / quick access)
        public string? LastDeviceName { get; set; }        // Last seen \\.\DISPLAYn
        public string? LastRegistryKey { get; set; }       // Last seen registry path (for copy)
        public string? LastSerialNumber { get; set; }      // Last seen EDID serial
        public string? LastInstanceId { get; set; }        // Last seen PnP instance id
        public string? LastMonitorId { get; set; }         // Last seen EDID model/product

        // Optional UI hints
        public int? PreferredOrder { get; set; }           // Optional pin for row ordering
        public int? LastKnownX { get; set; }               // Last known X position (for sorting)

        // NEW: known command targets we can try with MultiMonitorTool
        // e.g., "\"\\\\.\\DISPLAY2\"", "\"Generic PnP Monitor\""
        public List<string> KnownTargets { get; set; } = new List<string>();

        // NEW: user can mark this alias as the preferred primary monitor
        public bool IsPreferredPrimary { get; set; }
    }

    // A detected (or reconstructed) monitor instance used by the app
    public class DetectedMonitor
    {
        public string Name { get; set; } = string.Empty;        // NirSoft "Name" (sometimes includes device)
        public string DeviceName { get; set; } = string.Empty;  // \\.\DISPLAYn
        public string MonitorKey { get; set; } = string.Empty;  // Registry path if available
        public string MonitorId { get; set; } = string.Empty;   // EDID model/product
        public string InstanceId { get; set; } = string.Empty;  // PnP instance path
        public string SerialNumber { get; set; } = string.Empty;// EDID serial
        public string StableKey { get; set; } = string.Empty;   // Our persistent key
        public bool IsActive { get; set; }                      // Currently enabled/visible
        public bool IsPresent { get; set; }                     // Currently connected/present
        public int PositionX { get; set; }                      // X position for sorting
    }

    // Row used in the Settings grid
    public class AliasViewRow
    {
        public string StableKey { get; set; } = string.Empty;   // Full stable key (shown in tooltip)
        public string ShortKey { get; set; } = string.Empty;    // Abbreviated display in grid
        public string RegistryKey { get; set; } = string.Empty; // Copyable registry path
        public string? Alias { get; set; }                      // Editable alias
        public bool IsPreferredPrimary { get; set; }
    }
}
