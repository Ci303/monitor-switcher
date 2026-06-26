# Release Notes

## Unreleased

## v0.3.4 - 2026-06-26

### Added

- Added a per-monitor fallback primary setting used when disabling the current primary monitor.

### Improved

- Settings now opens at a content-aware size so the grid, details pane, and action buttons fit without manual resizing.
- Preferred primary is now re-applied after monitor enable and layout restore so saved-layout restoration cannot leave another monitor primary.
- Reduced duplication in settings and enable-monitor orchestration code.

## v0.3.3 - 2026-06-26

### Fixed

- Fixed duplicate phantom `DEV:` monitor rows appearing after disabling displays.
- Fixed inactive monitor enable targeting so the app prefers the current live inactive display identity before stale saved display aliases.
- Fixed disabling the current primary monitor by applying the display topology directly when MultiMonitorTool reports success but leaves the primary display unchanged.
- Fixed saved-layout restore after re-enabling monitors so primary display and positions are reapplied from the saved layout.
- Fixed notification-area icon initialisation so the tray icon is assigned before it is shown.

### Improved

- Added focused regression coverage for monitor target resolution, duplicate detected rows, disconnected-state presentation, and primary-disable topology positioning.

## v0.3.2 - 2026-05-26

### Fixed

- Fixed an intermittent tray-restore repaint issue where the main window could reopen with the toolbar and status text visible but the monitor rows blank.

## v0.3.1 - 2026-05-19

### Added

- Added single-instance startup behavior so launching Monitor Switcher again restores the existing window instead of opening a second copy.

## v0.3.0 - 2026-05-19

### Added

- Added named layout profiles with prompted **Save** and explicit **Restore** actions.
- Added notification-area support with a tray menu for open, refresh, save, restore, settings, and exit.
- Added Settings options for minimize-to-tray, start-with-Windows, and confirm-before-disable.
- Added a missing-`MultiMonitorTool` banner on the main window.
- Added a monitor identity details panel in Settings.
- Added a **Diagnostics** button that opens recent monitor action and layout profile events.

### Improved

- Automatic restore now uses the selected layout profile.
- Settings now shows shortened registry class paths in the grid while preserving full registry keys for tooltips, copying, details, and Regedit opening.
- Layout profile names are sanitized before storage to avoid invalid filename characters and profile-file collisions.
- Missing `MultiMonitorTool.exe` is reported in the main-window banner and diagnostics log instead of interrupting every startup with a modal warning.
- Stale saved `DEV:` display aliases are now merged into the current stable monitor entries instead of appearing as duplicate offline monitors.
- Release zips are now self-contained Windows x64 packages.
- Release workflow now has the permissions needed to attach generated release zips.
- Manual release workflow runs only attach assets to a GitHub release when a tag is supplied.

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
