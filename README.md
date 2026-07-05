# Monitor Switcher

Windows Forms utility for switching work monitor profiles with NirSoft `MultiMonitorTool`.

## What It Does

- Detects connected monitors and keeps stable aliases for them.
- Shows each monitor in a compact desktop UI with status, device hints, and quick enable/disable actions.
- Saves and restores monitor layouts, including named layout profiles.
- Lets you choose a preferred primary monitor, rename monitor aliases, toggle dark mode, and keep the main window always on top.
- Opens saved monitor registry keys from Settings for easier monitor troubleshooting and alias setup.
- Can minimize to the notification area and start with Windows.

## Requirements

- Windows 10 or Windows 11
- GitHub release zip: no separate .NET install required
- Local development: .NET 8 SDK
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
- Minimize-to-tray, start-with-Windows, restore-layout-on-app-start, and confirm-before-disable options.
- **Edit cfg** for `MultiMonitorTool.cfg`.
- **Open Registry** or double-clicking a registry-key cell to open Regedit at that monitor key. The grid shows a shortened class path, while tooltips/details and copy actions keep the full registry key.
- **Download MultiMonitorTool** for the NirSoft helper executable.
- **Update App** for downloading the latest GitHub release asset when one is published.
- A monitor identity details panel.
- A **Diagnostics** button that opens recent monitor action and layout profile events.

## Layout Profiles

Press **Save** to enter a layout profile name and capture the current monitor arrangement. Profile names are sanitized for safe file storage. Press **Restore** to load the currently selected profile. The active layout profile can be changed in Settings.

The app still auto-restores the selected layout profile after all saved monitors become active again.

If Windows or the graphics driver boots into the wrong rotation or position, enable both **Start with Windows** and **Restore layout on app start** in Settings. Monitor Switcher will then apply the selected layout profile shortly after sign-in.

## Running Locally

```powershell
dotnet build MonitorSwitcher.sln
dotnet run --project .\MonitorSwitcher\MonitorSwitcher.csproj
```

If you want the monitor actions to work during local development, copy the NirSoft files into `MonitorSwitcher\bin\Debug\net8.0-windows\` after build, or place them in the project directory before building so MSBuild copies them to the output folder.

## Releases

GitHub releases are packaged as a self-contained Windows x64 zip archive. There is currently no installer; download the release zip, extract it to a folder, and run `MonitorSwitcher.exe`.

If you want full monitor control features after extracting a release, use **Settings > Download MultiMonitorTool** or place `MultiMonitorTool.exe` and `MultiMonitorTool.cfg` beside `MonitorSwitcher.exe`.

## Layout Storage

The monitor layout snapshot is stored per user under:

```text
%APPDATA%\WorkMonitorSwitcher\monitor-layout.cfg
```

Named layout profiles are stored under:

```text
%APPDATA%\WorkMonitorSwitcher\layouts
```

Existing installs that previously used a layout file beside the executable are migrated automatically on first run.

## Repository Notes

- Main application code lives in `MonitorSwitcher`.
- User settings and monitor aliases are stored in `%APPDATA%\WorkMonitorSwitcher`.
- Publishing settings live under `MonitorSwitcher/Properties/PublishProfiles`.
- Third-party licensing notes are in `THIRD-PARTY-NOTICES.md`.
