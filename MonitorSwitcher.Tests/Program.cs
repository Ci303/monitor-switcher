using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WorkMonitorSwitcher;
using WorkMonitorSwitcher.Model;
using WorkMonitorSwitcher.Services;

var tests = new (string Name, Action Body)[]
{
    ("TargetsEquivalent treats quoted and DEV display names as the same device", TargetsEquivalentNormalisesDeviceNames),
    ("PruneKnownTargets removes display devices currently owned by another monitor", PruneKnownTargetsRemovesDeviceOwnedByOther),
    ("ResolveEnableTargetArgs rejects duplicate friendly names and stale devices", ResolveEnableTargetArgsRejectsAmbiguousTargets),
    ("ResolveEnableTargetArgs keeps a safe last-known display target", ResolveEnableTargetArgsKeepsSafeLastKnownDevice),
    ("ResolveEnableTargetArgs prefers live inactive device for the requested monitor", ResolveEnableTargetArgsPrefersLiveInactiveDevice),
    ("MonitorPresentationBuilder handles duplicate stable keys without throwing", PresentationBuilderHandlesDuplicateStableKeys),
    ("MonitorPresentationBuilder preserves disconnected detected state", PresentationBuilderPreservesDisconnectedState),
    ("DisplayTopologyService compacts remaining displays around the fallback primary", DisplayTopologyCompactsAroundFallbackPrimary),
    ("DisplayTopologyService maps saved orientation to CCD rotation", DisplayTopologyMapsSavedOrientationToCcdRotation),
    ("DisplayTopologyService matches saved layout by monitor identity after DISPLAY number changes", DisplayTopologyMatchesSavedLayoutByIdentity),
    ("AliasSettingsMapper applies aliases and primary selections", AliasSettingsMapperAppliesAliasesAndSelections),
    ("AliasSettingsMapper clears fallback when it matches preferred primary", AliasSettingsMapperClearsFallbackWhenItMatchesPreferredPrimary),
    ("PrimaryMonitorPreference resolves configured primary targets", PrimaryMonitorPreferenceResolvesConfiguredTargets),
    ("AtomicFileWriter replaces existing files and keeps a backup", AtomicFileWriterReplacesExistingFilesAndKeepsBackup),
    ("UiSettings disables automatic layout saves by default", UiSettingsDisablesAutomaticLayoutSavesByDefault),
    ("Form1 builds unique automatic layout backup paths", Form1BuildsUniqueAutomaticLayoutBackupPaths),
};

var failures = new List<string>();
foreach (var test in tests)
{
    try
    {
        test.Body();
        Console.WriteLine($"PASS {test.Name}");
    }
    catch (Exception ex)
    {
        failures.Add($"{test.Name}: {ex.Message}");
        Console.WriteLine($"FAIL {test.Name}");
        Console.WriteLine(ex);
    }
}

if (failures.Count > 0)
{
    Console.WriteLine();
    Console.WriteLine($"{failures.Count} test(s) failed:");
    foreach (var failure in failures)
        Console.WriteLine($"- {failure}");
    Environment.Exit(1);
}

Console.WriteLine();
Console.WriteLine($"{tests.Length} test(s) passed.");

static void TargetsEquivalentNormalisesDeviceNames()
{
    AssertTrue(
        MonitorTargetResolver.TargetsEquivalent("\"\\\\.\\DISPLAY1\"", "DEV:\\.\\DISPLAY1"),
        "Expected quoted display target and DEV fallback key to compare equal.");
}

static void PruneKnownTargetsRemovesDeviceOwnedByOther()
{
    var aliases = new Dictionary<string, MonitorInfo>(StringComparer.OrdinalIgnoreCase)
    {
        ["SN:RIGHT"] = new MonitorInfo
        {
            KnownTargets = new List<string>
            {
                "\"\\\\.\\DISPLAY1\"",
                "\\\\.\\DISPLAY3",
                "\\\\.\\DISPLAY1",
                "Right Friendly"
            }
        }
    };

    var detected = new List<DetectedMonitor>
    {
        Monitor("SN:LEFT", "\\\\.\\DISPLAY3", "Left Friendly", isActive: true),
        Monitor("SN:RIGHT", "\\\\.\\DISPLAY1", "Right Friendly", isActive: true)
    };

    MonitorTargetResolver.PruneKnownTargets(aliases, detected);

    var targets = aliases["SN:RIGHT"].KnownTargets;
    AssertFalse(targets.Any(t => MonitorTargetResolver.TargetsEquivalent(t, "\\\\.\\DISPLAY3")),
        "Expected DISPLAY3 to be pruned because it belongs to the left monitor.");
    AssertEquals(1, targets.Count(t => MonitorTargetResolver.TargetsEquivalent(t, "\\\\.\\DISPLAY1")),
        "Expected DISPLAY1 duplicates to be normalised to one target.");
    AssertContains(targets, "Right Friendly");
}

static void ResolveEnableTargetArgsRejectsAmbiguousTargets()
{
    var aliases = new Dictionary<string, MonitorInfo>(StringComparer.OrdinalIgnoreCase)
    {
        ["SN:RIGHT"] = new MonitorInfo
        {
            LastDeviceName = "\\\\.\\DISPLAY3",
            KnownTargets = new List<string>
            {
                "AOC Q27B3MA",
                "\\\\.\\DISPLAY3"
            }
        }
    };

    var detected = new List<DetectedMonitor>
    {
        Monitor("SN:LEFT", "\\\\.\\DISPLAY3", "AOC Q27B3MA", isActive: true),
        Monitor("SN:OTHER", "\\\\.\\DISPLAY2", "AOC Q27B3MA", isActive: true)
    };

    var targets = MonitorTargetResolver.ResolveEnableTargetArgs("SN:RIGHT", detected, aliases);

    AssertEquals(0, targets.Count, "Expected no safe enable targets when the name is duplicated and the device is in use.");
}

static void ResolveEnableTargetArgsKeepsSafeLastKnownDevice()
{
    var aliases = new Dictionary<string, MonitorInfo>(StringComparer.OrdinalIgnoreCase)
    {
        ["SN:RIGHT"] = new MonitorInfo
        {
            LastDeviceName = "\\\\.\\DISPLAY1",
            KnownTargets = new List<string>
            {
                "AOC Q27B3MA",
                "\\\\.\\DISPLAY3",
                "\\\\.\\DISPLAY1"
            }
        }
    };

    var detected = new List<DetectedMonitor>
    {
        Monitor("SN:LEFT", "\\\\.\\DISPLAY3", "AOC Q27B3MA", isActive: true)
    };

    var targets = MonitorTargetResolver.ResolveEnableTargetArgs("SN:RIGHT", detected, aliases);

    AssertSequence(targets, "\\\\.\\DISPLAY1");
}

static void ResolveEnableTargetArgsPrefersLiveInactiveDevice()
{
    var aliases = new Dictionary<string, MonitorInfo>(StringComparer.OrdinalIgnoreCase)
    {
        ["SN:RIGHT"] = new MonitorInfo
        {
            LastDeviceName = "\\\\.\\DISPLAY1",
            KnownTargets = new List<string>
            {
                "\\\\.\\DISPLAY1",
                "\\\\.\\DISPLAY8"
            }
        }
    };

    var detected = new List<DetectedMonitor>
    {
        Monitor("SN:LEFT", "\\\\.\\DISPLAY1", "AOC Q27B3MA", isActive: true),
        Monitor("SN:RIGHT", "\\\\.\\DISPLAY3", "AOC Q27B3MA", isActive: false)
    };

    var targets = MonitorTargetResolver.ResolveEnableTargetArgs("SN:RIGHT", detected, aliases);

    AssertSequence(targets, "\\\\.\\DISPLAY3", "\\\\.\\DISPLAY8");
}

static void PresentationBuilderHandlesDuplicateStableKeys()
{
    var detected = new List<DetectedMonitor>
    {
        new DetectedMonitor
        {
            StableKey = "SN:DUP",
            DeviceName = "\\\\.\\DISPLAY9",
            IsPresent = false,
            IsActive = false,
            PositionX = 100
        },
        new DetectedMonitor
        {
            StableKey = "SN:DUP",
            DeviceName = "\\\\.\\DISPLAY1",
            SerialNumber = "DUP",
            MonitorKey = "MK1",
            IsPresent = true,
            IsActive = true,
            PositionX = 0
        }
    };

    var rows = MonitorPresentationBuilder.Build(detected, EmptyAliases());

    AssertEquals(1, rows.Count, "Expected duplicate stable keys to collapse to one row.");
    AssertEquals("\\\\.\\DISPLAY1", rows[0].DeviceName, "Expected active complete row to win.");
    AssertTrue(rows[0].IsPresent, "Expected selected row to remain present.");
}

static void PresentationBuilderPreservesDisconnectedState()
{
    var detected = new List<DetectedMonitor>
    {
        Monitor("DEV:\\.\\DISPLAY4", "\\\\.\\DISPLAY4", "Phantom", isPresent: false, isActive: false)
    };

    var rows = MonitorPresentationBuilder.Build(detected, EmptyAliases());

    AssertEquals(1, rows.Count, "Expected detected row to be shown.");
    AssertFalse(rows[0].IsPresent, "Expected presentation builder not to force IsPresent to true.");
}

static void DisplayTopologyCompactsAroundFallbackPrimary()
{
    var left = new DisplaySourceLayout("\\\\.\\DISPLAY1", 1, -2560, 19, 2560, 1440);
    var right = new DisplaySourceLayout("\\\\.\\DISPLAY3", 3, 2560, -1087, 2560, 1440);

    var positions = DisplayTopologyService.CalculateCompactedPositions(
        new[] { left, right },
        left);

    AssertEquals(new DisplayPosition(0, 0), positions[1], "Expected fallback primary to move to origin.");
    AssertEquals(new DisplayPosition(2560, -1106), positions[3], "Expected right display to remain to the right with its vertical offset preserved.");
}

static void DisplayTopologyMapsSavedOrientationToCcdRotation()
{
    AssertTrue(DisplayTopologyService.TryMapDisplayOrientationToCcdRotation(0, out var identity),
        "Expected landscape orientation to map.");
    AssertTrue(DisplayTopologyService.TryMapDisplayOrientationToCcdRotation(1, out var rotate90),
        "Expected 90-degree orientation to map.");
    AssertTrue(DisplayTopologyService.TryMapDisplayOrientationToCcdRotation(2, out var rotate180),
        "Expected 180-degree orientation to map.");
    AssertTrue(DisplayTopologyService.TryMapDisplayOrientationToCcdRotation(3, out var rotate270),
        "Expected 270-degree orientation to map.");

    AssertEquals<uint>(1, identity, "Expected identity CCD rotation.");
    AssertEquals<uint>(2, rotate90, "Expected 90-degree CCD rotation.");
    AssertEquals<uint>(3, rotate180, "Expected 180-degree CCD rotation.");
    AssertEquals<uint>(4, rotate270, "Expected 270-degree CCD rotation.");
    AssertFalse(DisplayTopologyService.TryMapDisplayOrientationToCcdRotation(99, out _),
        "Expected unknown orientation to be rejected.");
}

static void DisplayTopologyMatchesSavedLayoutByIdentity()
{
    var dir = Path.Combine(Path.GetTempPath(), "MonitorSwitcher.Tests", Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(dir);

    try
    {
        var path = Path.Combine(dir, "monitor-layout.cfg");
        File.WriteAllText(path, string.Join(Environment.NewLine, new[]
        {
            "[Monitor0]",
            @"Name=\\.\DISPLAY1",
            @"MonitorID=MONITOR\AOC2703\{4d36e96e-e325-11ce-bfc1-08002be10318}\0000",
            "SerialNumber=17ZQ9HA007983",
            "Width=2560",
            "Height=1440",
            "DisplayOrientation=0",
            "PositionX=-2560",
            "PositionY=19",
            "[Monitor1]",
            @"Name=\\.\DISPLAY2",
            @"MonitorID=MONITOR\AOC2730\{4d36e96e-e325-11ce-bfc1-08002be10318}\0005",
            "SerialNumber=",
            "Width=2560",
            "Height=1440",
            "DisplayOrientation=0",
            "PositionX=0",
            "PositionY=0",
            "[Monitor2]",
            @"Name=\\.\DISPLAY3",
            @"MonitorID=MONITOR\AOC2703\{4d36e96e-e325-11ce-bfc1-08002be10318}\0004",
            "SerialNumber=17ZP6HA001814",
            "Width=1440",
            "Height=2560",
            "DisplayOrientation=3",
            "PositionX=2560",
            "PositionY=-1087"
        }));

        var detected = new[]
        {
            Monitor("SN:17ZQ9HA007983", @"\\.\DISPLAY3", "Left", isActive: true),
            Monitor("MK:MIDDLE", @"\\.\DISPLAY1", "Middle", isActive: true),
            Monitor("SN:17ZP6HA001814", @"\\.\DISPLAY2", "Right", isActive: true)
        };
        detected[0].SerialNumber = "17ZQ9HA007983";
        detected[1].MonitorId = @"MONITOR\AOC2730\{4d36e96e-e325-11ce-bfc1-08002be10318}\0005";
        detected[2].SerialNumber = "17ZP6HA001814";

        var map = DisplayTopologyService.ResolveSavedLayoutDeviceNameMap(
            path,
            detected,
            new[] { @"\\.\DISPLAY1", @"\\.\DISPLAY2", @"\\.\DISPLAY3" });

        AssertEquals(@"\\.\DISPLAY2", map[@"\\.\DISPLAY1"], "Expected current middle display to use the saved middle layout.");
        AssertEquals(@"\\.\DISPLAY3", map[@"\\.\DISPLAY2"], "Expected current right display to use the saved right layout.");
        AssertEquals(@"\\.\DISPLAY1", map[@"\\.\DISPLAY3"], "Expected current left display to use the saved left layout.");
    }
    finally
    {
        if (Directory.Exists(dir))
            Directory.Delete(dir, recursive: true);
    }
}

static void AliasSettingsMapperAppliesAliasesAndSelections()
{
    var aliases = new Dictionary<string, MonitorInfo>(StringComparer.OrdinalIgnoreCase)
    {
        ["SN:LEFT"] = new MonitorInfo { Name = "Old Left", IsPreferredPrimary = true },
        ["SN:RIGHT"] = new MonitorInfo { Name = "Old Right" },
        ["SN:STALE"] = new MonitorInfo { Name = "Stale" }
    };

    var updated = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["SN:LEFT"] = "Left Monitor",
        ["SN:RIGHT"] = "Right Monitor"
    };

    AliasSettingsMapper.ApplyMonitorSettings(
        aliases,
        new[] { "SN:STALE" },
        updated,
        preferredPrimaryKey: "SN:RIGHT",
        fallbackPrimaryKey: "SN:LEFT");

    AssertFalse(aliases.ContainsKey("SN:STALE"), "Expected removed monitor key to be deleted.");
    AssertEquals("Left Monitor", aliases["SN:LEFT"].Name, "Expected alias to be updated.");
    AssertFalse(aliases["SN:LEFT"].IsPreferredPrimary, "Expected previous preferred primary to be cleared.");
    AssertTrue(aliases["SN:RIGHT"].IsPreferredPrimary, "Expected selected preferred primary to be set.");
    AssertTrue(aliases["SN:LEFT"].IsFallbackPrimary, "Expected selected fallback primary to be set.");
    AssertFalse(aliases["SN:RIGHT"].IsFallbackPrimary, "Expected fallback primary to remain single-select.");

    var rows = AliasSettingsMapper.BuildRows(
        new[] { Monitor("SN:LEFT", "\\\\.\\DISPLAY1", "AOC", isActive: true) },
        aliases);

    AssertEquals("Left Monitor", rows[0].Alias, "Expected settings row to use saved alias.");
    AssertTrue(rows[0].IsFallbackPrimary, "Expected settings row to expose fallback selection.");
}

static void AliasSettingsMapperClearsFallbackWhenItMatchesPreferredPrimary()
{
    var aliases = new Dictionary<string, MonitorInfo>(StringComparer.OrdinalIgnoreCase)
    {
        ["SN:LEFT"] = new MonitorInfo { Name = "Left Monitor" }
    };

    AliasSettingsMapper.ApplyMonitorSettings(
        aliases,
        Array.Empty<string>(),
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
        preferredPrimaryKey: "SN:LEFT",
        fallbackPrimaryKey: "SN:LEFT");

    AssertTrue(aliases["SN:LEFT"].IsPreferredPrimary, "Expected preferred primary to be set.");
    AssertFalse(aliases["SN:LEFT"].IsFallbackPrimary, "Expected same-key fallback primary to be cleared.");
}

static void PrimaryMonitorPreferenceResolvesConfiguredTargets()
{
    var aliases = new Dictionary<string, MonitorInfo>(StringComparer.OrdinalIgnoreCase)
    {
        ["SN:LEFT"] = new MonitorInfo { IsFallbackPrimary = true },
        ["SN:RIGHT"] = new MonitorInfo { IsPreferredPrimary = true }
    };

    var detected = new List<DetectedMonitor>
    {
        new DetectedMonitor
        {
            StableKey = "SN:LEFT",
            DeviceName = "\\\\.\\DISPLAY1",
            Name = "Left",
            IsPresent = true,
            IsActive = true,
            PositionX = -2560
        },
        new DetectedMonitor
        {
            StableKey = "SN:RIGHT",
            DeviceName = "\\\\.\\DISPLAY3",
            Name = "Right",
            IsPresent = true,
            IsActive = true,
            PositionX = 2560
        }
    };

    var preferred = PrimaryMonitorPreference.ResolvePreferredPrimaryTarget(detected, aliases);
    var fallback = PrimaryMonitorPreference.ResolveConfiguredFallbackDeviceName("\\\\.\\DISPLAY3", detected, aliases);
    var excludedFallback = PrimaryMonitorPreference.ResolveConfiguredFallbackDeviceName("\\\\.\\DISPLAY1", detected, aliases);
    var automatic = PrimaryMonitorPreference.ResolveLeftMostActiveTarget(detected, aliases);

    AssertEquals("\\\\.\\DISPLAY3", preferred, "Expected active preferred primary target to resolve.");
    AssertEquals("\\\\.\\DISPLAY1", fallback, "Expected configured fallback device to resolve.");
    AssertEquals<string?>(null, excludedFallback, "Expected excluded fallback device to be ignored.");
    AssertEquals("\\\\.\\DISPLAY1", automatic, "Expected left-most active monitor to be automatic fallback.");
}

static void AtomicFileWriterReplacesExistingFilesAndKeepsBackup()
{
    var dir = Path.Combine(Path.GetTempPath(), "MonitorSwitcher.Tests", Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(dir);

    try
    {
        var path = Path.Combine(dir, "settings.json");
        File.WriteAllText(path, "old");

        AtomicFileWriter.WriteAllText(path, "new");

        AssertEquals("new", File.ReadAllText(path), "Expected target file to contain replacement content.");
        AssertTrue(File.Exists(path + ".bak"), "Expected a backup file to be created for replaced content.");
        AssertEquals("old", File.ReadAllText(path + ".bak"), "Expected backup file to contain previous content.");
    }
    finally
    {
        if (Directory.Exists(dir))
            Directory.Delete(dir, recursive: true);
    }
}

static void UiSettingsDisablesAutomaticLayoutSavesByDefault()
{
    var settings = new UiSettings();

    AssertFalse(settings.AutoSaveLayoutBeforeDisable,
        "Expected automatic layout saves before disable to be off by default.");
}

static void Form1BuildsUniqueAutomaticLayoutBackupPaths()
{
    var dir = Path.Combine(Path.GetTempPath(), "MonitorSwitcher.Tests", Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(dir);

    try
    {
        var layoutPath = Path.Combine(dir, "monitor-layout.cfg");
        var timestamp = new DateTime(2026, 7, 2, 21, 16, 22);

        var first = Form1.NextAutoSaveBackupPath(layoutPath, timestamp);
        AssertEquals(
            layoutPath + ".autosave-20260702-211622.bak",
            first,
            "Expected first automatic backup path to use the timestamp.");

        File.WriteAllText(first, "existing backup");
        var second = Form1.NextAutoSaveBackupPath(layoutPath, timestamp);
        AssertEquals(
            layoutPath + ".autosave-20260702-211622-2.bak",
            second,
            "Expected backup path to avoid overwriting an existing backup.");
    }
    finally
    {
        if (Directory.Exists(dir))
            Directory.Delete(dir, recursive: true);
    }
}

static DetectedMonitor Monitor(
    string stableKey,
    string deviceName,
    string name,
    bool isActive,
    bool isPresent = true)
{
    return new DetectedMonitor
    {
        StableKey = stableKey,
        DeviceName = deviceName,
        Name = name,
        IsActive = isActive,
        IsPresent = isPresent
    };
}

static Dictionary<string, MonitorInfo> EmptyAliases()
    => new(StringComparer.OrdinalIgnoreCase);

static void AssertTrue(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}

static void AssertFalse(bool condition, string message)
{
    if (condition) throw new InvalidOperationException(message);
}

static void AssertEquals<T>(T expected, T actual, string message)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"{message} Expected '{expected}', got '{actual}'.");
}

static void AssertContains(IEnumerable<string> values, string expected)
{
    if (!values.Any(v => string.Equals(v, expected, StringComparison.OrdinalIgnoreCase)))
        throw new InvalidOperationException($"Expected collection to contain '{expected}'.");
}

static void AssertSequence(IReadOnlyList<string> actual, params string[] expected)
{
    if (actual.Count != expected.Length)
        throw new InvalidOperationException($"Expected {expected.Length} item(s), got {actual.Count}: {string.Join(", ", actual)}.");

    for (int i = 0; i < expected.Length; i++)
    {
        if (!string.Equals(actual[i], expected[i], StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException($"Expected item {i} to be '{expected[i]}', got '{actual[i]}'.");
    }
}
