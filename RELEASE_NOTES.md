# Release Notes

## v0.2.0

### Added

- Refreshed main monitor switcher UI with clearer monitor rows, status badges, device hints, and hover tooltips.
- Added a monitor summary to the main window showing active, present, and saved monitor counts.
- Added a settings-window title-bar icon that matches the main app.
- Added **Open Registry** in Settings.
- Added support for double-clicking a **Registry Key** cell to open Regedit at that monitor key.
- Added tooltips for settings and monitor actions.

### Improved

- Renamed the main layout action to **Save Layout** for clarity.
- Improved Settings layout with clearer grouping for monitor tools versus save/cancel/remove actions.
- Improved Settings grid sizing, row presentation, and dark-theme rendering.
- Documented automatic monitor detection, settings tools, and where **Download MultiMonitorTool** installs the helper.

### Validation

- Built successfully with `dotnet build MonitorSwitcher.sln -c Release`.
