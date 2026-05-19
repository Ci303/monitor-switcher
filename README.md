# Monitor Switcher

Windows Forms utility for switching work monitor profiles with NirSoft `MultiMonitorTool`.

## What It Does

- Detects connected monitors and keeps stable aliases for them.
- Shows each monitor in a compact desktop UI with status, device hints, and quick enable/disable actions.
- Saves and restores a full monitor layout when all displays are active again.
- Lets you choose a preferred primary monitor, rename monitor aliases, toggle dark mode, and keep the main window always on top.
- Opens saved monitor registry keys from Settings for easier monitor troubleshooting and alias setup.

## Requirements

- Windows 10 or Windows 11
- .NET 8 SDK for local development
- NirSoft `MultiMonitorTool`

## Third-Party Tool Setup

Download `MultiMonitorTool` from the official NirSoft page and place these files beside the built app executable.
Settings includes a **Download MultiMonitorTool** button that downloads `MultiMonitorTool.exe` to the same folder the app is running from, which is the location Monitor Switcher uses at runtime:

- `MultiMonitorTool.exe`
- `MultiMonitorTool.cfg`

Official source:

- https://www.nirsoft.net/utils/multi_monitor_tool.html

The app will still start without the tool and can fall back to Windows' active-screen detection, but enable/disable/primary/layout actions and richer monitor identity work best when the tool is present.

## Detection

Monitor Switcher detects displays automatically when the app starts and refreshes when Windows reports display, device, or settings changes. You can also force a scan with **Refresh**.

Detection order:

1. `MultiMonitorTool.exe` CSV export.
2. `MultiMonitorTool.exe` text export.
3. Windows `Screen.AllScreens` fallback.

## Settings

Settings includes:

- Editable monitor aliases.
- Preferred primary monitor selection.
- Dark mode and always-on-top toggles.
- **Edit cfg** for `MultiMonitorTool.cfg`.
- **Open Registry** or double-clicking a registry-key cell to open Regedit at that monitor key.
- **Download MultiMonitorTool** for the NirSoft helper executable.
- **Update App** for downloading the latest GitHub release asset when one is published.

## Running Locally

```powershell
dotnet build MonitorSwitcher.sln
dotnet run --project .\MonitorSwitcher\MonitorSwitcher.csproj
```

If you want the monitor actions to work during local development, copy the NirSoft files into `MonitorSwitcher\bin\Debug\net8.0-windows\` after build, or place them in the project directory before building so MSBuild copies them to the output folder.

## Releases

GitHub releases are packaged as a zip archive of the published app output. There is currently no installer; download the release zip, extract it to a folder, and run `MonitorSwitcher.exe`.

If you want full monitor control features after extracting a release, use **Settings > Download MultiMonitorTool** or place `MultiMonitorTool.exe` and `MultiMonitorTool.cfg` beside `MonitorSwitcher.exe`.

## Layout Storage

The monitor layout snapshot is stored per user under:

```text
%APPDATA%\WorkMonitorSwitcher\monitor-layout.cfg
```

Existing installs that previously used a layout file beside the executable are migrated automatically on first run.

## Repository Notes

- Main application code lives in `MonitorSwitcher`.
- User settings and monitor aliases are stored in `%APPDATA%\WorkMonitorSwitcher`.
- Publishing settings live under `MonitorSwitcher/Properties/PublishProfiles`.
- Third-party licensing notes are in `THIRD-PARTY-NOTICES.md`.
