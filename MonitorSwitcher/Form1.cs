#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

using WorkMonitorSwitcher.Model;
using WorkMonitorSwitcher.Services;
using WorkMonitorSwitcher.UI;
using System.Threading.Tasks;


namespace WorkMonitorSwitcher
{
    public partial class Form1 : Form
    {
        // ---- External tool + config paths ----
        private readonly string _toolPath = Path.Combine(AppContext.BaseDirectory, "MultiMonitorTool.exe");
        private readonly string _bundledLayoutPath = Path.Combine(AppContext.BaseDirectory, "monitor-layout.cfg");
        private readonly string _layoutPath;

        // Per-user app data dir
        private readonly string _appDataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "WorkMonitorSwitcher");

        // Stores (JSON) and services
        private readonly UiSettingsStore _uiStore;
        private readonly AliasStore _aliasStore;
        private readonly DetectionService _detectSvc;
        private readonly LayoutService _layoutSvc;

        // In-memory state
        private UiSettings _uiSettings;
        private readonly Dictionary<string, MonitorInfo> _aliasMap = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, MonitorControls> _controlsByKey = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<Control> _dynamicControls = new();
        private List<DetectedMonitor> _detected = new();

        // Debounced refresh on display changes
        private readonly System.Windows.Forms.Timer _refreshTimer = new() { Interval = 800 };

        // ---- Layout constants / handles ----
        private const int SideMargin = 12;
        private const int ControlGapX = 10;
        private const int ButtonWidth = 120;
        private const int ButtonHeight = 30;
        private const int RowVerticalGap = 80;
        private const int TopButtonsY = 10;
        private const int FirstRowY = 60; // fallback start if no HR yet

        private Button? _btnSettings;
        private Button? _btnRefresh;
        private Button? _btnSaveLayout;
        private Panel? _hrTop;

        private int _layoutRightMost;

        // Tracks rows currently executing a tool command
        private readonly HashSet<string> _busy = new(StringComparer.OrdinalIgnoreCase);

        // Restore robustness
        private bool _deferredLayout; // schedule a full rebuild after restore/show
        private bool IsNormalVisible => Visible && WindowState == FormWindowState.Normal;

        private sealed class ToolResult
        {
            public int ExitCode { get; init; }
            public string StdOut { get; init; } = string.Empty;
            public string StdErr { get; init; } = string.Empty;
        }

        public Form1()
        {
            // Reduce flicker
            SetStyle(
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.UserPaint |
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.ResizeRedraw,  // repaint background on resize
            true);

            UpdateStyles();

            InitializeComponent();

            Text = "Monitor Switcher";
            AutoSize = false;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimumSize = new Size(420, 260); // avoids restore-to-titlebar

            Directory.CreateDirectory(_appDataDir);
            _layoutPath = Path.Combine(_appDataDir, "monitor-layout.cfg");
            EnsureUserLayoutPath();

            _uiStore = new UiSettingsStore(_appDataDir);
            _aliasStore = new AliasStore(_appDataDir);
            _detectSvc = new DetectionService(_toolPath);
            _layoutSvc = new LayoutService(_toolPath);

            _uiSettings = _uiStore.LoadOrDefault();
            var aliases = _aliasStore.Load();
            foreach (var kv in aliases) _aliasMap[kv.Key] = kv.Value;

            // Top buttons
            _btnSettings = new Button { Text = "Settings", Size = new Size(80, 30) };
            _btnSettings.Click += SettingsButton_Click;
            Controls.Add(_btnSettings);

            _btnRefresh = new Button { Text = "Refresh", Size = new Size(80, 30) };
            _btnRefresh.Click += (_, __) => RefreshMonitorsAndUi();
            Controls.Add(_btnRefresh);

            _btnSaveLayout = new Button { Text = "Save", Size = new Size(80, 30) };
            _btnSaveLayout.Click += (_, __) =>
            {
                var ok = _layoutSvc.SaveLayout(_layoutPath);
                if (!ok)
                    ThemedMessageBox.Info(this,
                        "Unable to save layout. Is MultiMonitorTool.exe present?",
                        "Save Layout", _uiSettings.DarkMode);
                else
                    ThemedMessageBox.Info(this, "Layout saved.", "Save Layout", _uiSettings.DarkMode);
            };
            Controls.Add(_btnSaveLayout);

            _hrTop = new Panel { Height = 1, BackColor = SystemColors.ControlDark, Visible = true };
            Controls.Add(_hrTop);

            SizeChanged += (_, __) => LeftTopButtons();
            Layout += (_, __) => LeftTopButtons();

            CheckToolExists(); // warning only; app still runs

            // Apply theme and caption styling
            ApplyTheme(_uiSettings.DarkMode);

            // Honor persisted "always on top" (reinforced again in OnShown)
            TopMost = _uiSettings.AlwaysOnTop;

            // Restore window position/size (safe-guarded)
            RestoreWindowBounds();

            _refreshTimer.Tick += (_, __) => { _refreshTimer.Stop(); RefreshMonitorsAndUi(); };

            LeftTopButtons();
            UpdateTopSeparator();

            RefreshMonitorsAndUi();

            FormClosing += (_, __) =>
            {
                SaveWindowBounds();
                _uiStore.Save(_uiSettings);
            };
        }

        // Re-assert TopMost and complete any deferred layout when first shown
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            BeginInvoke(new Action(() =>
            {
                try { TopMost = _uiSettings.AlwaysOnTop; } catch { /* ignore */ }

                if (_deferredLayout)
                {
                    _deferredLayout = false;
                    RefreshMonitorsAndUi();
                    LeftTopButtons();
                }
            }));
        }

        // Finish deferred layout on restore from minimized
        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);

            if (IsNormalVisible && _deferredLayout)
            {
                _deferredLayout = false;
                RefreshMonitorsAndUi();
                LeftTopButtons();
            }
        }

        // Composited painting for children (extra flicker reduction)
        protected override CreateParams CreateParams
        {
            get
            {
                const int WS_EX_COMPOSITED = 0x02000000;
                var cp = base.CreateParams;
                cp.ExStyle |= WS_EX_COMPOSITED;
                return cp;
            }
        }
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            // Explicitly clear the entire client area to our BackColor to avoid dark-mode “black box” artifacts
            e.Graphics.Clear(this.BackColor);
        }

        // React to display topology changes
        protected override void WndProc(ref Message m)
        {
            const int WM_DISPLAYCHANGE = 0x007E;
            const int WM_DEVICECHANGE = 0x0219;
            const int WM_SETTINGCHANGE = 0x001A;

            if (m.Msg == WM_DISPLAYCHANGE || m.Msg == WM_DEVICECHANGE || m.Msg == WM_SETTINGCHANGE)
            {
                _refreshTimer.Stop();
                _refreshTimer.Start();
            }
            base.WndProc(ref m);
        }

        // ---------- Guardrails ----------
        // Which Windows.Forms.Screen is currently hosting this window?
        private Screen GetHostScreen()
        {
            try { return Screen.FromRectangle(this.Bounds); }
            catch { return Screen.PrimaryScreen ?? Screen.AllScreens.First(); }
        }

        // Try to get a live device name (\\.\DISPLAYn) for a given stable key.
        // Returns null if not currently present.
        private string? TryGetDeviceNameForStableKey(string stableKey)
        {
            var live = _detected.FirstOrDefault(d =>
                d.IsPresent &&
                d.StableKey.Equals(stableKey, StringComparison.OrdinalIgnoreCase));

            return string.IsNullOrWhiteSpace(live?.DeviceName) ? null : live!.DeviceName;
        }

        // From a device name (\\.\DISPLAYn), get the corresponding Screen (if any)
        private Screen? TryGetScreenForDevice(string deviceName)
        {
            if (string.IsNullOrWhiteSpace(deviceName)) return null;
            return Screen.AllScreens.FirstOrDefault(s =>
                string.Equals(s.DeviceName, deviceName, StringComparison.OrdinalIgnoreCase));
        }

        // Choose a safe fallback active screen to host the window,
        // excluding the screen identified by excludeDevice (the one we’re about to disable).
        // Prefers the left-most active screen; falls back to Primary; then any.
        private Screen GetFallbackActiveScreen(string? excludeDevice)
        {
            // 1) Active detected monitors mapped to Screen
            var activeDevices = _detected
                .Where(d => d.IsPresent && d.IsActive)
                .Select(d => d.DeviceName)
                .Where(dn => !string.IsNullOrWhiteSpace(dn) &&
                             !string.Equals(dn, excludeDevice, StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var dev in activeDevices.OrderBy(d => d, StringComparer.OrdinalIgnoreCase))
            {
                var scr = TryGetScreenForDevice(dev);
                if (scr != null) return scr;
            }

            // 2) Primary screen, unless it’s the excluded one
            var primary = Screen.PrimaryScreen;
            if (primary != null &&
                !string.Equals(primary.DeviceName, excludeDevice, StringComparison.OrdinalIgnoreCase))
                return primary;

            // 3) Any other screen not excluded
            var any = Screen.AllScreens.FirstOrDefault(s =>
                !string.Equals(s.DeviceName, excludeDevice, StringComparison.OrdinalIgnoreCase));
            return any ?? (primary ?? Screen.AllScreens.First());
        }
        // Center this window on the given screen's working area.
        // Temporarily drops TopMost to avoid z-order oddities while moving.
        private void RehomeWindowToScreen(Screen target)
        {
            var wa = target.WorkingArea;

            // Keep a sensible minimum size; clamp to the target work area.
            int w = Math.Min(wa.Width, Math.Max(420, this.Width));
            int h = Math.Min(wa.Height, Math.Max(260, this.Height));

            int x = wa.Left + Math.Max(0, (wa.Width - w) / 2);
            int y = wa.Top + Math.Max(0, (wa.Height - h) / 2);

            var prevTopMost = this.TopMost;
            try
            {
                this.TopMost = false;                 // prevent flicker/z-fight while moving
                this.StartPosition = FormStartPosition.Manual;
                this.Bounds = new Rectangle(x, y, w, h);
                this.Activate();                      // bring to front on its new screen
            }
            finally
            {
                this.TopMost = prevTopMost;          // restore user's preference
            }
        }

        // If this window currently lives on the screen identified by 'disablingDevice',
        // move it to a safe, still-active screen before we disable that display.
        private void MoveWindowIfHostedOn(string? disablingDevice)
        {
            if (string.IsNullOrWhiteSpace(disablingDevice))
                return;

            var host = GetHostScreen();
            if (!string.Equals(host.DeviceName, disablingDevice, StringComparison.OrdinalIgnoreCase))
                return; // we're not on the soon-to-be-disabled display

            var fallback = GetFallbackActiveScreen(disablingDevice);
            RehomeWindowToScreen(fallback);
        }

        // ---------- Top layout helpers ----------

        private void LeftTopButtons()
        {
            if (_btnSettings == null || _btnRefresh == null || _btnSaveLayout == null) return;

            // Don't lay out while minimized; mark that we owe a layout later.
            if (!IsNormalVisible) { _deferredLayout = true; return; }

            const int spacing = 12;
            int x = SideMargin;

            _btnSettings.Location = new Point(x, TopButtonsY);
            x = _btnSettings.Right + spacing;

            _btnRefresh.Location = new Point(x, TopButtonsY);
            x = _btnRefresh.Right + spacing;

            _btnSaveLayout.Location = new Point(x, TopButtonsY);

            _btnSettings.BringToFront();
            _btnRefresh.BringToFront();
            _btnSaveLayout.BringToFront();

            UpdateTopSeparator();
        }

        private int GetSeparatorY()
        {
            int buttonHeight = _btnSettings?.Height ?? 30;
            return TopButtonsY + buttonHeight + 10;
        }

        private void UpdateTopSeparator()
        {
            if (_hrTop == null) return;
            int y = GetSeparatorY();
            _hrTop.Location = new Point(SideMargin, y);
            _hrTop.Width = Math.Max(0, ClientSize.Width - (SideMargin * 2));
            _hrTop.Height = 1;
            // color is updated by ApplyTheme
        }

        // ---------- Theme / caption ----------

        private void ApplyTheme(bool dark)
        {
            var palette = dark ? ThemePalette.Dark() : ThemePalette.Light();

            // Apply to this form + children
            Themer.Apply(this, palette);

            // Keep HR visible
            if (_hrTop != null) _hrTop.BackColor = palette.Border;

            // Native title bar
            DwmInterop.SetDarkTitleBar(this.Handle, dark);
            var accent = WindowsTheme.AccentColor();
            if (accent.HasValue) DwmInterop.SetCaptionColor(this.Handle, accent.Value);

            Invalidate(true);
            Update();
        }

        // ---------- Settings dialog ----------
        private void SettingsButton_Click(object? sender, EventArgs e)
        {
            var rows = BuildPresentationList()
                .Select(d =>
                {
                    var alias = GetAliasFor(d.StableKey);
                    _aliasMap.TryGetValue(d.StableKey, out var mi);
                    var reg = mi?.LastRegistryKey ?? string.Empty;
                    return new AliasViewRow
                    {
                        StableKey = d.StableKey,
                        ShortKey = d.StableKey.Length <= 28 ? d.StableKey : "…" + d.StableKey[^28..],
                        Alias = alias,
                        RegistryKey = reg,
                        IsPreferredPrimary = (mi?.IsPreferredPrimary ?? false)
                    };
                })
                .ToList();

            using var dlg = new AliasSettingsForm(rows, _uiSettings.DarkMode, _uiSettings.AlwaysOnTop)
            {
                StartPosition = FormStartPosition.CenterParent,
                ShowInTaskbar = false,
                TopMost = true
            };

            var prevTopMost = this.TopMost;
            DialogResult dr;
            try
            {
                // Ensure the dialog can float above everything without fighting the host
                this.TopMost = false;
                dlg.TopMost = true;

                dr = dlg.ShowDialog(this);
            }
            finally
            {
                // Restore whatever the user had configured for the main window
                this.TopMost = prevTopMost;
            }

            if (dr == DialogResult.OK)
            {
                // Aliases
                foreach (var kvp in dlg.UpdatedMappings)
                {
                    if (!_aliasMap.TryGetValue(kvp.Key, out var info))
                        info = new MonitorInfo();
                    info.Name = kvp.Value ?? info.Name;
                    _aliasMap[kvp.Key] = info;
                }

                // Preferred primary (single selection)
                foreach (var kv in _aliasMap.Keys.ToList())
                {
                    var info = _aliasMap[kv];
                    info.IsPreferredPrimary = string.Equals(kv, dlg.PreferredPrimaryKey, StringComparison.OrdinalIgnoreCase);
                    _aliasMap[kv] = info;
                }

                _aliasStore.Save(_aliasMap);

                // Dark mode
                if (dlg.DarkModeResult.HasValue && dlg.DarkModeResult.Value != _uiSettings.DarkMode)
                {
                    _uiSettings.DarkMode = dlg.DarkModeResult.Value;
                    _uiStore.Save(_uiSettings);
                    ApplyTheme(_uiSettings.DarkMode);
                }

                // Always on top
                if (dlg.AlwaysOnTopResult.HasValue && dlg.AlwaysOnTopResult.Value != _uiSettings.AlwaysOnTop)
                {
                    _uiSettings.AlwaysOnTop = dlg.AlwaysOnTopResult.Value;
                    _uiStore.Save(_uiSettings);
                    TopMost = _uiSettings.AlwaysOnTop;
                }

                RefreshMonitorsAndUi();
            }
        }

        // ---------- Working change ----------
        private void SetRowBusy(string stableKey, bool busy)
        {
            if (!_controlsByKey.TryGetValue(stableKey, out var ctrls))
                return;

            ctrls.IsBusy = busy;

            // While busy: show WORKING… and disable both buttons
            if (busy)
            {
                ctrls.StatusLabel.Text = "WORKING…";
                var palette = _uiSettings.DarkMode ? ThemePalette.Dark() : ThemePalette.Light();
                ctrls.StatusLabel.ForeColor = palette.StatusBusy;

                ctrls.DisableButton.Enabled = false;
                ctrls.EnableButton.Enabled = false;
                var paletteBtn = _uiSettings.DarkMode ? ThemePalette.Dark() : ThemePalette.Light();
                ctrls.DisableButton.ForeColor = paletteBtn.TextSubtle;
                ctrls.EnableButton.ForeColor = paletteBtn.TextSubtle;
            }
            else
            {
                // When we clear busy, we’ll immediately re-run UpdateButtonStatus()
                // to put the correct ONLINE/DISABLED/OFFLINE text and button states back.
            }

            // Keep the status aligned over the Enable button (same math used elsewhere)
            var statusSize = ctrls.StatusLabel.PreferredSize;
            int statusX = ctrls.EnableButton.Right - statusSize.Width;
            int statusY = ctrls.TitleLabel.Top + Math.Max(0, (ctrls.TitleLabel.Height - statusSize.Height) / 2);
            ctrls.StatusLabel.Location = new Point(statusX, statusY);

            ctrls.StatusLabel.Invalidate();
        }

        // ---------- Refresh + dynamic UI build ----------

        private void RefreshMonitorsAndUi()
        {
            LeftTopButtons();
            UpdateTopSeparator();

            _detected = _detectSvc.Detect();

            // Attempt to re-bind aliases if stable keys changed (e.g., port swaps)
            ReconcileAliasesForDetected();
            ReconcilePositionalAliases();

            // First-run seeding if exactly three monitors and no aliases
            if (_aliasMap.Count == 0 && _detected.Count == 3)
            {
                var ordered = _detected.OrderBy(d => d.PositionX).ToList();
                SetAlias(ordered[0].StableKey, "Left Monitor");
                SetAlias(ordered[1].StableKey, "Middle Monitor");
                SetAlias(ordered[2].StableKey, "Right Monitor");
            }

            // Keep alias metadata up to date + remember targets
            foreach (var m in _detected)
            {
                if (!_aliasMap.TryGetValue(m.StableKey, out var info))
                {
                    info = new MonitorInfo();
                    _aliasMap[m.StableKey] = info;
                }
                if (!string.IsNullOrWhiteSpace(m.DeviceName))
                    info.LastDeviceName = m.DeviceName;
                info.LastKnownX = m.PositionX;
                if (!string.IsNullOrWhiteSpace(m.MonitorKey))
                    info.LastRegistryKey = m.MonitorKey;
                if (!string.IsNullOrWhiteSpace(m.SerialNumber))
                    info.LastSerialNumber = m.SerialNumber;
                if (!string.IsNullOrWhiteSpace(m.InstanceId))
                    info.LastInstanceId = m.InstanceId;
                if (!string.IsNullOrWhiteSpace(m.MonitorId))
                    info.LastMonitorId = m.MonitorId;

                // NEW: remember useful command targets (device + friendly name)
                EnsureKnownTargets(m.StableKey, m.DeviceName, m.Name);
            }
            _aliasStore.Save(_aliasMap);

            var toShow = BuildPresentationList();

            SuspendLayout();
            try
            {
                foreach (var c in _dynamicControls) { Controls.Remove(c); c.Dispose(); }
                _dynamicControls.Clear();
                _controlsByKey.Clear();

                _layoutRightMost = 0;

                int sepY = GetSeparatorY();
                int yStart = sepY + (_hrTop?.Height ?? 1) + 12;
                int y = Math.Max(FirstRowY, yStart);

                foreach (var m in toShow)
                {
                    AddMonitorControls(m, GetAliasFor(m.StableKey), y);
                    y += RowVerticalGap;
                }

                int desiredWidth = Math.Max(320, _layoutRightMost + SideMargin);

                // Also ensure there’s room for the top buttons row
                int topButtonsRight = (_btnSaveLayout?.Right ?? 0);
                desiredWidth = Math.Max(desiredWidth, topButtonsRight + SideMargin);

                int desiredHeight = Math.Max(120, y + SideMargin);

                if (IsNormalVisible)
                {
                    ClientSize = new Size(desiredWidth, desiredHeight);
                }
                else
                {
                    _deferredLayout = true; // resize later
                }
                var palette = _uiSettings.DarkMode ? ThemePalette.Dark() : ThemePalette.Light();
                Themer.Apply(this, palette);
                if (_hrTop != null) _hrTop.BackColor = palette.Border;


                UpdateButtonStatus();
            }
            finally
            {
                ResumeLayout(performLayout: true);
            }
        }

        private List<DetectedMonitor> BuildPresentationList()
        {
            var presentByKey = _detected.ToDictionary(d => d.StableKey, StringComparer.OrdinalIgnoreCase);
            var allKeys = new HashSet<string>(presentByKey.Keys, StringComparer.OrdinalIgnoreCase);
            foreach (var key in _aliasMap.Keys) allKeys.Add(key);

            var list = new List<DetectedMonitor>();
            foreach (var key in allKeys)
            {
                if (presentByKey.TryGetValue(key, out var d))
                {
                    d.IsPresent = true;
                    list.Add(d);
                }
                else
                {
                    var info = _aliasMap.TryGetValue(key, out var mi) ? mi : new MonitorInfo();
                    list.Add(new DetectedMonitor
                    {
                        StableKey = key,
                        Name = GetAliasFor(key),
                        DeviceName = info.LastDeviceName ?? string.Empty,
                        MonitorKey = info.LastRegistryKey ?? string.Empty,
                        MonitorId = string.Empty,
                        InstanceId = string.Empty,
                        SerialNumber = string.Empty,
                        IsActive = false,
                        IsPresent = false,
                        PositionX = info.LastKnownX ?? 0
                    });
                }
            }

            return list
                .OrderBy(m => GetPreferredOrder(m.StableKey))
                .ThenBy(m => AliasHintRank(GetAliasFor(m.StableKey)))
                .ThenBy(m => m.PositionX)
                .ThenBy(m => GetAliasFor(m.StableKey), StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private int GetPreferredOrder(string stableKey)
        {
            if (_aliasMap.TryGetValue(stableKey, out var info) && info.PreferredOrder.HasValue)
                return info.PreferredOrder.Value;
            return int.MaxValue - 100000;
        }

        private static int AliasHintRank(string alias)
        {
            if (alias.Contains("Left", StringComparison.OrdinalIgnoreCase)) return 0;
            if (alias.Contains("Middle", StringComparison.OrdinalIgnoreCase)) return 1;
            if (alias.Contains("Centre", StringComparison.OrdinalIgnoreCase)) return 1;
            if (alias.Contains("Right", StringComparison.OrdinalIgnoreCase)) return 2;
            return 50;
        }

        private void AddMonitorControls(DetectedMonitor monitor, string friendlyName, int positionY)
        {
            int labelX = SideMargin;
            int disableX = SideMargin;
            int enableX = disableX + ButtonWidth + ControlGapX;

            var label = new Label
            {
                Text = friendlyName,
                Location = new Point(labelX, positionY),
                AutoSize = true,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            Controls.Add(label);
            _dynamicControls.Add(label);

            var statusLabel = new Label { AutoSize = true };
            Controls.Add(statusLabel);
            _dynamicControls.Add(statusLabel);

            var buttonOff = new Button
            {
                Text = "DISABLE",
                Location = new Point(disableX, positionY + 25),
                Size = new Size(ButtonWidth, ButtonHeight),
                Tag = monitor.StableKey
            };
            buttonOff.Click += ButtonOff_Click;
            Controls.Add(buttonOff);
            _dynamicControls.Add(buttonOff);

            var buttonOn = new Button
            {
                Text = "ENABLE",
                Location = new Point(enableX, positionY + 25),
                Size = new Size(ButtonWidth, ButtonHeight),
                Tag = monitor.StableKey
            };
            buttonOn.Click += ButtonOn_Click;
            Controls.Add(buttonOn);
            _dynamicControls.Add(buttonOn);

            // status text over the Enable button (right-aligned), vertically aligned with the title label
            var palette = _uiSettings.DarkMode ? ThemePalette.Dark() : ThemePalette.Light();

            statusLabel.Text = monitor.IsPresent ? (monitor.IsActive ? "ONLINE" : "DISABLED") : "OFFLINE";
            statusLabel.ForeColor = monitor.IsPresent ? (monitor.IsActive ? palette.StatusOk : palette.StatusWarn) : palette.TextSubtle;

            var statusSize = statusLabel.PreferredSize;
            int statusX = buttonOn.Right - statusSize.Width;
            int statusY = label.Top + Math.Max(0, (label.Height - statusSize.Height) / 2);
            statusLabel.Location = new Point(statusX, statusY);

            _layoutRightMost = Math.Max(_layoutRightMost, Math.Max(buttonOn.Right, statusLabel.Right));

            _controlsByKey[monitor.StableKey] = new MonitorControls
            {
                DisableButton = buttonOff,
                EnableButton = buttonOn,
                StatusLabel = statusLabel,
                TitleLabel = label
            };
        }

        // ---------- Button handlers ----------

        private async void ButtonOff_Click(object? sender, EventArgs e)
        {
            if (sender is not Button btn || btn.Tag is not string stableKey) return;

            var present = _detected.Where(d => d.IsPresent).ToList();
            var activeCount = present.Count(d => d.IsActive);
            if (activeCount <= 1)
            {
                ThemedMessageBox.Warn(this,
                    "You cannot disable this monitor — at least one monitor must remain active.",
                    "Monitor Switcher", _uiSettings.DarkMode);
                return;
            }

            var target = ResolveDisableTargetArg(stableKey);
            if (string.IsNullOrWhiteSpace(target))
            {
                LogMonitorAction($"DISABLE {stableKey}: no target resolved.");
                ThemedMessageBox.Info(this,
                    "Cannot disable: no resolvable device identifier for this monitor.",
                    "Monitor Switcher", _uiSettings.DarkMode);
                return;
            }
            LogMonitorAction($"DISABLE {stableKey}: selected target '{target}'.");

            // --- NEW: if this window is currently on the display we're about to disable,
            // move it to a safe, still-active screen first.
            // We try to get the live \\.\DISPLAYn for this stableKey.
            var disablingDevice = TryGetDeviceNameForStableKey(stableKey);
            if (!string.IsNullOrWhiteSpace(disablingDevice))
            {
                MoveWindowIfHostedOn(disablingDevice);
                // tiny settle helps avoid a visible jump if the desktop is busy
                await Task.Delay(100);
            }

            // Capture baseline layout if everything is currently active
            bool allActiveNow = _detected.Count > 0 && _detected.All(d => d.IsPresent && d.IsActive);
            if (allActiveNow) _layoutSvc.SaveLayout(_layoutPath);

            // Show WORKING… and lock the row
            SetRowBusy(stableKey, true);
            UpdateButtonStatus();

            var res = await ExecToolAsync("/disable", target);
            LogMonitorAction($"DISABLE {stableKey}: exit={res.ExitCode}, stderr='{res.StdErr}'.");

            // Give the desktop a brief moment to settle
            await Task.Delay(400);

            // Clear busy and reflect result
            SetRowBusy(stableKey, false);

            if (res.ExitCode != 0)
            {
                ThemedMessageBox.Error(this,
                    $"Disable failed (exit {res.ExitCode})." +
                    (string.IsNullOrWhiteSpace(res.StdErr) ? string.Empty : $"\n{res.StdErr}"),
                    "Monitor Switcher", _uiSettings.DarkMode);
            }

            RefreshMonitorsAndUi();
        }


        private async void ButtonOn_Click(object? sender, EventArgs e)
        {
            if (sender is not Button btn || btn.Tag is not string stableKey) return;

            var enableTargets = ResolveEnableTargetArgs(stableKey).ToList();
            if (enableTargets.Count == 0)
            {
                LogMonitorAction($"ENABLE {stableKey}: no targets resolved.");
                ThemedMessageBox.Info(this,
                    "Cannot enable: no resolvable monitor identifiers for this monitor. Toggle another display on, then press Refresh.",
                    "Monitor Switcher", _uiSettings.DarkMode);
                return;
            }
            LogMonitorAction($"ENABLE {stableKey}: candidate targets [{string.Join(", ", enableTargets.Select(t => $"'{t}'"))}].");

            // Show WORKING… and lock the row
            SetRowBusy(stableKey, true);
            UpdateButtonStatus();

            ToolResult? last = null;
            bool enabledTarget = false;
            foreach (var target in enableTargets)
            {
                LogMonitorAction($"ENABLE {stableKey}: trying target '{target}'.");
                last = await ExecToolAsync("/enable", target);
                LogMonitorAction($"ENABLE {stableKey}: target '{target}' exit={last.ExitCode}, stderr='{last.StdErr}'.");
                await Task.Delay(500);

                var detectedNow = _detectSvc.Detect();
                if (IsStableKeyActive(detectedNow, stableKey))
                {
                    LogMonitorAction($"ENABLE {stableKey}: activation confirmed after target '{target}'.");
                    enabledTarget = true;
                    break;
                }

                LogMonitorAction($"ENABLE {stableKey}: target '{target}' did not activate requested stable key.");
            }

            // Clear busy
            SetRowBusy(stableKey, false);

            if (!enabledTarget)
            {
                LogMonitorAction($"ENABLE {stableKey}: all target attempts failed.");
                ThemedMessageBox.Error(this,
                    $"Enable failed to activate the selected monitor." +
                    (last == null ? string.Empty : $" (last exit {last.ExitCode}).") +
                    (string.IsNullOrWhiteSpace(last?.StdErr) ? string.Empty : $"\n{last!.StdErr}"),
                    "Monitor Switcher", _uiSettings.DarkMode);

                RefreshMonitorsAndUi();
                return;
            }

            // Success path mirrors your previous logic
            RefreshMonitorsAndUi();
            EnforcePrimaryMonitorOrder();

            // If all are now active, restore saved layout
            if (_detected.Count > 0 && _detected.All(d => d.IsPresent && d.IsActive))
            {
                if (!_layoutSvc.LoadLayout(_layoutPath))
                {
                    // Optional: notify if missing
                    // ThemedMessageBox.Warn(this, "Saved monitor layout file not found.", "Monitor Switcher", _uiSettings.DarkMode);
                }

                await Task.Delay(800);
                RefreshMonitorsAndUi();
            }
        }

        private void UpdateButtonStatus()
        {
            var palette = _uiSettings.DarkMode ? ThemePalette.Dark() : ThemePalette.Light();

            var present = _detected.Where(d => d.IsPresent).ToList();
            int activeCount = present.Count(d => d.IsActive);

            foreach (var kv in _controlsByKey)
            {
                var key = kv.Key;
                var ctrls = kv.Value;

                if (ctrls.IsBusy)
                {
                    ctrls.StatusLabel.Text = "WORKING…";
                    ctrls.StatusLabel.ForeColor = palette.StatusBusy;

                    ctrls.DisableButton.Enabled = false;
                    ctrls.EnableButton.Enabled = false;
                    ctrls.DisableButton.ForeColor = palette.TextSubtle;
                    ctrls.EnableButton.ForeColor = palette.TextSubtle;

                    var sizeBusy = ctrls.StatusLabel.PreferredSize;
                    int sxBusy = ctrls.EnableButton.Right - sizeBusy.Width;
                    int syBusy = ctrls.TitleLabel.Top + Math.Max(0, (ctrls.TitleLabel.Height - sizeBusy.Height) / 2);
                    ctrls.StatusLabel.Location = new Point(sxBusy, syBusy);

                    continue;
                }

                var live = present.FirstOrDefault(d => d.StableKey.Equals(key, StringComparison.OrdinalIgnoreCase));
                var info = _aliasMap.TryGetValue(key, out var mi) ? mi : null;

                bool isPresent = live != null;
                bool isActive = live?.IsActive == true;
                bool hasLast = !string.IsNullOrWhiteSpace(info?.LastDeviceName);
                bool hasKnownTargets = info != null && info.KnownTargets.Any();

                ctrls.StatusLabel.Text = isPresent ? (isActive ? "ONLINE" : "DISABLED") : "OFFLINE";
                ctrls.StatusLabel.ForeColor = isPresent
                    ? (isActive ? palette.StatusOk : palette.StatusWarn)
                    : palette.TextSubtle;

                var statusSize = ctrls.StatusLabel.PreferredSize;
                int statusX = ctrls.EnableButton.Right - statusSize.Width;
                int statusY = ctrls.TitleLabel.Top + Math.Max(0, (ctrls.TitleLabel.Height - statusSize.Height) / 2);
                ctrls.StatusLabel.Location = new Point(statusX, statusY);

                bool canDisable = isPresent && isActive && activeCount > 1;
                bool canEnable = !isActive && (hasLast || hasKnownTargets);

                ctrls.DisableButton.Enabled = canDisable;
                ctrls.DisableButton.ForeColor = canDisable ? ForeColor : palette.TextSubtle;

                ctrls.EnableButton.Enabled = canEnable;
                ctrls.EnableButton.ForeColor = canEnable ? ForeColor : palette.TextSubtle;
            }
        }


        private void EnforcePrimaryMonitorOrder()
        {
            // Prefer the user's selected Preferred Primary; else left-most active
            var primary = _aliasMap.FirstOrDefault(kv => kv.Value.IsPreferredPrimary);
            if (!string.IsNullOrEmpty(primary.Key))
            {
                var m = _detected.FirstOrDefault(d => d.IsPresent && d.IsActive &&
                                                      d.StableKey.Equals(primary.Key, StringComparison.OrdinalIgnoreCase));
                if (m != null)
                {
                    var target = ResolveTargetArg(primary.Key);
                    if (!string.IsNullOrWhiteSpace(target))
                    {
                        _layoutSvc.SetPrimary(target);
                        return;
                    }
                }
            }

            var firstActive = _detected.Where(d => d.IsPresent && d.IsActive)
                                       .OrderBy(d => d.PositionX)
                                       .FirstOrDefault();
            if (firstActive != null)
            {
                var target = ResolveTargetArg(firstActive.StableKey);
                if (!string.IsNullOrWhiteSpace(target))
                    _layoutSvc.SetPrimary(target);
            }
        }

        // ---------- Helpers ----------

        private void ReconcileAliasesForDetected()
        {
            if (_aliasMap.Count == 0 || _detected.Count == 0) return;

            var detectedKeys = new HashSet<string>(_detected.Select(d => d.StableKey), StringComparer.OrdinalIgnoreCase);
            var missing = _detected.Where(d =>
                !_aliasMap.ContainsKey(d.StableKey) ||
                string.IsNullOrWhiteSpace(_aliasMap[d.StableKey].Name)).ToList();
            if (missing.Count == 0) return;

            foreach (var m in missing)
            {
                if (_aliasMap.TryGetValue(m.StableKey, out var existing) &&
                    !string.IsNullOrWhiteSpace(existing.Name))
                    continue;

                var match = FindUniqueAliasMatch(m, detectedKeys);
                if (match == null) continue;

                var oldKey = match.Value.Key;
                var info = match.Value.Value;

                _aliasMap.Remove(oldKey);
                _aliasMap[m.StableKey] = info;
            }
        }

        private KeyValuePair<string, MonitorInfo>? FindUniqueAliasMatch(DetectedMonitor m, HashSet<string> detectedKeys)
        {
            // Only consider aliases that are NOT currently detected, to avoid swaps.
            IEnumerable<KeyValuePair<string, MonitorInfo>> candidates = _aliasMap
                .Where(kv => !detectedKeys.Contains(kv.Key));

            KeyValuePair<string, MonitorInfo>? TryMatch(Func<MonitorInfo, string?> selector, string? value)
            {
                if (string.IsNullOrWhiteSpace(value)) return null;
                var matches = candidates
                    .Where(kv => StringsEqual(selector(kv.Value), value))
                    .ToList();
                return matches.Count == 1 ? matches[0] : null;
            }

            var match = TryMatch(i => i.LastSerialNumber, m.SerialNumber);
            if (match != null) return match;

            match = TryMatch(i => i.LastInstanceId, m.InstanceId);
            if (match != null) return match;

            match = TryMatch(i => i.LastRegistryKey, m.MonitorKey);
            if (match != null) return match;

            match = TryMatch(i => i.LastMonitorId, m.MonitorId);
            if (match != null) return match;

            var model = GetMonitorModelKey(m.MonitorId);
            if (!string.IsNullOrWhiteSpace(model))
            {
                var modelMatches = candidates
                    .Where(kv => string.Equals(GetMonitorModelKey(kv.Value.LastMonitorId), model, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (modelMatches.Count == 1)
                    return modelMatches[0];

                if (modelMatches.Count > 1)
                {
                    var byX = modelMatches
                        .Where(kv => kv.Value.LastKnownX.HasValue)
                        .Select(kv => new
                        {
                            Candidate = kv,
                            Delta = Math.Abs((kv.Value.LastKnownX ?? 0) - m.PositionX)
                        })
                        .OrderBy(x => x.Delta)
                        .ToList();

                    if (byX.Count == 1)
                        return byX[0].Candidate;

                    if (byX.Count > 1 && byX[0].Delta < byX[1].Delta)
                        return byX[0].Candidate;
                }
            }

            return null;
        }

        private static string GetMonitorModelKey(string? monitorId)
        {
            if (string.IsNullOrWhiteSpace(monitorId)) return string.Empty;
            var t = monitorId.Trim();

            // Typical format: MONITOR\AOC2730\{guid}\0001 -> keep MONITOR\AOC2730
            var parts = t.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2)
                return $"{parts[0]}\\{parts[1]}";

            return t;
        }

        private void ReconcilePositionalAliases()
        {
            if (_aliasMap.Count == 0 || _detected.Count == 0) return;

            var present = _detected.Where(d => d.IsPresent).OrderBy(d => d.PositionX).ToList();
            if (present.Count < 3) return;

            var leftKv = _aliasMap.FirstOrDefault(kv => kv.Value.Name.Equals("Left Monitor", StringComparison.OrdinalIgnoreCase));
            var middleKv = _aliasMap.FirstOrDefault(kv => kv.Value.Name.Equals("Middle Monitor", StringComparison.OrdinalIgnoreCase));
            var rightKv = _aliasMap.FirstOrDefault(kv => kv.Value.Name.Equals("Right Monitor", StringComparison.OrdinalIgnoreCase));

            if (string.IsNullOrWhiteSpace(leftKv.Key) ||
                string.IsNullOrWhiteSpace(middleKv.Key) ||
                string.IsNullOrWhiteSpace(rightKv.Key))
                return;

            var presentKeys = new HashSet<string>(present.Select(d => d.StableKey), StringComparer.OrdinalIgnoreCase);
            int lmrPresentCount = 0;
            if (presentKeys.Contains(leftKv.Key)) lmrPresentCount++;
            if (presentKeys.Contains(middleKv.Key)) lmrPresentCount++;
            if (presentKeys.Contains(rightKv.Key)) lmrPresentCount++;

            // Only intervene if at least one positional alias is currently orphaned.
            if (lmrPresentCount == 3) return;

            var desired = new[]
            {
                (Name: "Left Monitor", Key: present[0].StableKey),
                (Name: "Middle Monitor", Key: present[1].StableKey),
                (Name: "Right Monitor", Key: present[2].StableKey)
            };

            foreach (var d in desired)
            {
                var currentNameOwner = _aliasMap.FirstOrDefault(kv =>
                    kv.Value.Name.Equals(d.Name, StringComparison.OrdinalIgnoreCase));
                if (string.IsNullOrWhiteSpace(currentNameOwner.Key)) continue;
                if (currentNameOwner.Key.Equals(d.Key, StringComparison.OrdinalIgnoreCase)) continue;

                if (!_aliasMap.TryGetValue(d.Key, out var targetInfo))
                {
                    targetInfo = new MonitorInfo();
                    _aliasMap[d.Key] = targetInfo;
                }

                var displacedName = targetInfo.Name;
                targetInfo.Name = d.Name;
                currentNameOwner.Value.Name = displacedName;
            }
        }

        private static bool StringsEqual(string? a, string? b)
        {
            if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b)) return false;
            return string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private string GetAliasFor(string stableKey)
        {
            return _aliasMap.TryGetValue(stableKey, out var info) && !string.IsNullOrWhiteSpace(info.Name)
                ? info.Name
                : stableKey;
        }

        private void SetAlias(string stableKey, string alias)
        {
            if (!_aliasMap.TryGetValue(stableKey, out var info))
                info = new MonitorInfo();
            info.Name = alias;
            _aliasMap[stableKey] = info;
        }

        // New helpers for robust target resolution
        private static string NormaliseTarget(string s)
        {
            var t = (s ?? string.Empty).Trim();
            if (t.Length >= 2 &&
                t.StartsWith("\"", StringComparison.Ordinal) &&
                t.EndsWith("\"", StringComparison.Ordinal))
            {
                t = t[1..^1];
            }
            return t;
        }

        private void EnsureKnownTargets(string stableKey, params string?[] candidates)
        {
            if (!_aliasMap.TryGetValue(stableKey, out var info))
            {
                info = new MonitorInfo();
                _aliasMap[stableKey] = info;
            }

            for (int i = candidates.Length - 1; i >= 0; i--)
            {
                var c = candidates[i];
                var t = (c ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(t)) continue;

                t = NormaliseTarget(t);

                int existingIndex = info.KnownTargets.FindIndex(x => string.Equals(x, t, StringComparison.OrdinalIgnoreCase));
                if (existingIndex >= 0)
                    info.KnownTargets.RemoveAt(existingIndex);

                // Most-recent targets go first to avoid stale DISPLAYn mappings.
                info.KnownTargets.Insert(0, t);
            }

            const int MaxKnownTargets = 8;
            if (info.KnownTargets.Count > MaxKnownTargets)
                info.KnownTargets.RemoveRange(MaxKnownTargets, info.KnownTargets.Count - MaxKnownTargets);
        }

        private string? ResolveTargetArg(string stableKey)
        {
            // Fall back to live detection
            var live = _detected.FirstOrDefault(d => d.StableKey.Equals(stableKey, StringComparison.OrdinalIgnoreCase));
            if (live != null)
            {
                var liveDevice = NormaliseTarget(live.DeviceName);
                if (!string.IsNullOrWhiteSpace(liveDevice)) return liveDevice;

                var liveName = NormaliseTarget(live.Name);
                if (!string.IsNullOrWhiteSpace(liveName) && IsUniqueMonitorName(liveName))
                    return liveName;
            }

            // Prefer last-known device name before older, potentially stale targets
            if (_aliasMap.TryGetValue(stableKey, out var info))
            {
                if (!string.IsNullOrWhiteSpace(info.LastDeviceName) &&
                    !IsDeviceNameInUseByOther(stableKey, info.LastDeviceName))
                    return NormaliseTarget(info.LastDeviceName);

                var knownDevice = info.KnownTargets
                    .FirstOrDefault(t => IsLikelyDeviceName(t) && !IsDeviceNameInUseByOther(stableKey, t));
                if (!string.IsNullOrWhiteSpace(knownDevice))
                    return knownDevice;

                var anyKnown = info.KnownTargets.FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));
                if (!string.IsNullOrWhiteSpace(anyKnown))
                    return anyKnown;
            }

            return null;
        }

        private string? ResolveDisableTargetArg(string stableKey)
        {
            var live = _detected.FirstOrDefault(d =>
                d.StableKey.Equals(stableKey, StringComparison.OrdinalIgnoreCase) &&
                d.IsPresent &&
                d.IsActive);
            if (live != null)
            {
                var liveDevice = NormaliseTarget(live.DeviceName);
                if (!string.IsNullOrWhiteSpace(liveDevice))
                    return liveDevice;
            }

            return ResolveTargetArg(stableKey);
        }

        private IEnumerable<string> ResolveEnableTargetArgs(string stableKey)
        {
            var candidates = new List<string>();
            void Add(string? raw)
            {
                var t = NormaliseTarget(raw ?? string.Empty);
                if (string.IsNullOrWhiteSpace(t)) return;
                if (candidates.Any(x => x.Equals(t, StringComparison.OrdinalIgnoreCase))) return;
                candidates.Add(t);
            }

            var live = _detected.FirstOrDefault(d => d.StableKey.Equals(stableKey, StringComparison.OrdinalIgnoreCase));
            if (live != null)
            {
                var liveName = NormaliseTarget(live.Name);
                if (!string.IsNullOrWhiteSpace(liveName) && !IsLikelyDeviceName(liveName) && IsUniqueMonitorName(liveName))
                    Add(liveName);
            }

            if (_aliasMap.TryGetValue(stableKey, out var info))
            {
                foreach (var knownName in info.KnownTargets.Where(t => !IsLikelyDeviceName(t)))
                {
                    if (IsUniqueMonitorName(knownName))
                        Add(knownName);
                }

                if (!string.IsNullOrWhiteSpace(info.LastDeviceName) &&
                    !IsDeviceNameInUseByOther(stableKey, info.LastDeviceName))
                {
                    Add(info.LastDeviceName);
                }

                foreach (var knownDevice in info.KnownTargets.Where(t =>
                             IsLikelyDeviceName(t) && !IsDeviceNameInUseByOther(stableKey, t)))
                {
                    Add(knownDevice);
                }

                foreach (var knownAny in info.KnownTargets)
                    Add(knownAny);
            }

            Add(ResolveTargetArg(stableKey));
            return candidates;
        }

        private static bool IsStableKeyActive(IEnumerable<DetectedMonitor> detected, string stableKey)
        {
            return detected.Any(d =>
                d.StableKey.Equals(stableKey, StringComparison.OrdinalIgnoreCase) &&
                d.IsPresent &&
                d.IsActive);
        }

        private static bool IsLikelyDeviceName(string? t)
        {
            if (string.IsNullOrWhiteSpace(t)) return false;
            var v = t.Trim();
            return v.StartsWith(@"\\.\DISPLAY", StringComparison.OrdinalIgnoreCase);
        }

        private bool IsDeviceNameInUseByOther(string stableKey, string? deviceName)
        {
            if (string.IsNullOrWhiteSpace(deviceName)) return false;
            var live = _detected.FirstOrDefault(d =>
                d.DeviceName.Equals(deviceName, StringComparison.OrdinalIgnoreCase));
            return live != null && !live.StableKey.Equals(stableKey, StringComparison.OrdinalIgnoreCase);
        }

        private bool IsUniqueMonitorName(string name)
        {
            int count = _detected.Count(d => d.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            return count == 1;
        }

        private static void LogMonitorAction(string message)
        {
            Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
        }

        private void CheckToolExists()
        {
            if (!File.Exists(_toolPath))
            {
                ThemedMessageBox.Warn(this,
                    $"MultiMonitorTool.exe not found at: {_toolPath}\n" +
                    "Download MultiMonitorTool from https://www.nirsoft.net/utils/multi_monitor_tool.html and place it beside the app executable.\n" +
                    "Enable/disable/primary/layout actions will be unavailable. Detection will fall back to Screen.AllScreens.",
                    "Monitor Switcher", _uiSettings.DarkMode);
            }
        }

        private static readonly TimeSpan ToolTimeout = TimeSpan.FromSeconds(10);

        private void EnsureUserLayoutPath()
        {
            if (File.Exists(_layoutPath))
                return;

            try
            {
                if (File.Exists(_bundledLayoutPath))
                    File.Copy(_bundledLayoutPath, _layoutPath, overwrite: false);
            }
            catch
            {
                // Non-fatal: a layout file will be created on the first successful save.
            }
        }

        private async Task<ToolResult> ExecToolAsync(params string[] args)
        {
            if (!File.Exists(_toolPath))
            {
                return new ToolResult
                {
                    ExitCode = -3,
                    StdOut = string.Empty,
                    StdErr = $"Tool not found: {_toolPath}"
                };
            }

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = _toolPath,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true
                };

                foreach (var a in args)
                {
                    var t = NormaliseTarget(a ?? string.Empty);
                    if (string.IsNullOrWhiteSpace(t)) continue;
                    psi.ArgumentList.Add(t);
                }

                using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };

                process.Start();

                var stdOutTask = process.StandardOutput.ReadToEndAsync();
                var stdErrTask = process.StandardError.ReadToEndAsync();

                using var cts = new CancellationTokenSource(ToolTimeout);
                try
                {
                    await process.WaitForExitAsync(cts.Token);
                }
                catch (OperationCanceledException)
                {
                    try
                    {
                        if (!process.HasExited)
                            process.Kill(entireProcessTree: true);
                    }
                    catch { /* ignore */ }

                    return new ToolResult
                    {
                        ExitCode = -2,
                        StdOut = string.Empty,
                        StdErr = $"Timed out after {ToolTimeout.TotalSeconds:0}s."
                    };
                }

                return new ToolResult
                {
                    ExitCode = process.ExitCode,
                    StdOut = await stdOutTask,
                    StdErr = await stdErrTask
                };
            }
            catch (Exception ex)
            {
                ThemedMessageBox.Error(this,
                    $"Error executing tool: {ex.Message}",
                    "Monitor Switcher", _uiSettings.DarkMode);

                return new ToolResult
                {
                    ExitCode = -1,
                    StdOut = string.Empty,
                    StdErr = ex.Message
                };
            }
        }


        // ---------- Window bounds persistence ----------

        private void RestoreWindowBounds()
        {
            if (_uiSettings.WindowWidth > 0 && _uiSettings.WindowHeight > 0)
            {
                StartPosition = FormStartPosition.Manual;
                var rect = new Rectangle(_uiSettings.WindowX, _uiSettings.WindowY, _uiSettings.WindowWidth, _uiSettings.WindowHeight);

                // Ensure at least partially on-screen
                var isOnScreen = Screen.AllScreens.Any(s => s.WorkingArea.IntersectsWith(rect));
                if (!isOnScreen)
                {
                    StartPosition = FormStartPosition.CenterScreen;
                    return;
                }
                Bounds = rect;
            }
        }

        private void SaveWindowBounds()
        {
            if (WindowState == FormWindowState.Normal)
            {
                _uiSettings.WindowX = Bounds.X;
                _uiSettings.WindowY = Bounds.Y;
                _uiSettings.WindowWidth = Bounds.Width;
                _uiSettings.WindowHeight = Bounds.Height;
            }
        }
    }

    // Small UI handle bag (no logic)
    public class MonitorControls
    {
        public Button DisableButton { get; set; } = new();
        public Button EnableButton { get; set; } = new();
        public Label StatusLabel { get; set; } = new();
        public Label TitleLabel { get; set; } = new();
        public bool IsBusy { get; set; }
    }
}
