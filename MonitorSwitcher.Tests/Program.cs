using System;
using System.Collections.Generic;
using System.Linq;
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
    ("AliasSettingsMapper applies aliases and primary selections", AliasSettingsMapperAppliesAliasesAndSelections),
    ("AliasSettingsMapper clears fallback when it matches preferred primary", AliasSettingsMapperClearsFallbackWhenItMatchesPreferredPrimary),
    ("PrimaryMonitorPreference resolves configured primary targets", PrimaryMonitorPreferenceResolvesConfiguredTargets),
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
