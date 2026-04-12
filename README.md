# Monitor Switcher

Windows Forms utility for switching work monitor profiles with NirSoft `MultiMonitorTool`.

## What It Does

- Detects connected monitors and keeps stable aliases for them.
- Enables and disables individual displays from a small desktop UI.
- Saves and restores a full monitor layout when all displays are active again.
- Lets you choose a preferred primary monitor and basic UI settings.

## Requirements

- Windows 10 or Windows 11
- .NET 8 SDK for local development
- NirSoft `MultiMonitorTool`

## Third-Party Tool Setup

This repository no longer redistributes NirSoft binaries. Download `MultiMonitorTool` from the official NirSoft page and place these files beside the built app executable:

- `MultiMonitorTool.exe`
- `MultiMonitorTool.cfg`

Official source:

- https://www.nirsoft.net/utils/multi_monitor_tool.html

The app will still start without the tool, but enable/disable/primary/layout actions will be unavailable until the files are present.

## Running Locally

```powershell
dotnet build MonitorSwitcher.sln
dotnet run --project .\MonitorSwitcher\MonitorSwitcher.csproj
```

If you want the monitor actions to work during local development, copy the NirSoft files into `MonitorSwitcher\bin\Debug\net8.0-windows\` after build, or place them in the project directory before building so MSBuild copies them to the output folder.

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
