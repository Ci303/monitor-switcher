#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

using WorkMonitorSwitcher.Model;
using WorkMonitorSwitcher.Services;
using WorkMonitorSwitcher.UI;
using System.Threading.Tasks;
using static WorkMonitorSwitcher.Services.MonitorTargetResolver;


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
        private readonly LayoutProfileStore _profileStore;
        private readonly DiagnosticsLog _log;
        private readonly DetectionService _detectSvc;
        private readonly LayoutService _layoutSvc;
        private readonly DisplayTopologyService _topologySvc;

        // In-memory state
        private UiSettings _uiSettings;
        private readonly Dictionary<string, MonitorInfo> _aliasMap = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, MonitorControls> _controlsByKey = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<Control> _dynamicControls = new();
        private List<DetectedMonitor> _detected = new();
        private string _lastDetectionLogSignature = string.Empty;

        // Debounced refresh on display changes
        private readonly System.Windows.Forms.Timer _refreshTimer = new() { Interval = 800 };

        // ---- Layout constants / handles ----
        private const int SideMargin = 14;
        private const int ControlGapX = 10;
        private const int ButtonWidth = 108;
        private const int ButtonHeight = 30;
        private const int RowPanelWidth = 390;
        private const int RowPanelHeight = 82;
        private const int RowVerticalGap = 92;
        private const int TopButtonsY = 12;
        private const int FirstRowY = 68; // fallback start if no HR yet

        private Button? _btnSettings;
        private Button? _btnRefresh;
        private Button? _btnSaveLayout;
        private Button? _btnRestoreLayout;
        private Label? _summaryLabel;
        private Panel? _hrTop;
        private Panel? _missingToolPanel;
        private Label? _missingToolLabel;
        private Button? _missingToolSettingsButton;
        private NotifyIcon? _trayIcon;
        private ContextMenuStrip? _trayMenu;
        private readonly ToolTip _toolTip = new() { InitialDelay = 350, ReshowDelay = 100, AutoPopDelay = 8000 };

        private int _layoutRightMost;

        // Tracks rows currently executing a tool command
        private readonly HashSet<string> _busy = new(StringComparer.OrdinalIgnoreCase);
        private readonly SemaphoreSlim _displayActionGate = new(1, 1);
        private bool _exitRequested;
        private bool _restoringFromTray;
        private bool _isRefreshingUi;
        private bool _refreshAgainAfterCurrent;
        private System.Windows.Forms.Timer? _restoreRefreshTimer;
        private EventWaitHandle? _showExistingWindowEvent;
        private RegisteredWaitHandle? _showExistingWindowWait;

        // Restore robustness
        private bool _deferredLayout; // schedule a full rebuild after restore/show
        private bool IsNormalVisible => Visible && WindowState == FormWindowState.Normal;

        private sealed class ToolResult
        {
            public int ExitCode { get; init; }
            public string StdOut { get; init; } = string.Empty;
            public string StdErr { get; init; } = string.Empty;
        }

        private sealed class PrimaryDisableAttempt
        {
            public bool Attempted { get; init; }
            public bool Success { get; init; }
            public string ErrorMessage { get; init; } = string.Empty;
        }

        private sealed class EnableAttemptResult
        {
            public bool Success { get; init; }
            public ToolResult? LastResult { get; init; }
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
            _profileStore = new LayoutProfileStore(_appDataDir, _layoutPath);
            _log = new DiagnosticsLog(_appDataDir);
            _detectSvc = new DetectionService(_toolPath, _log.Write);
            _layoutSvc = new LayoutService(_toolPath);
            _topologySvc = new DisplayTopologyService();

            _uiSettings = _uiStore.LoadOrDefault();
            var startupEnabled = StartupManager.IsEnabled();
            if (_uiSettings.StartWithWindows && !startupEnabled)
            {
                StartupManager.SetEnabled(true, Application.ExecutablePath);
                startupEnabled = StartupManager.IsEnabled();
            }

            if (_uiSettings.StartWithWindows != startupEnabled)
            {
                _uiSettings.StartWithWindows = startupEnabled;
                _uiStore.Save(_uiSettings);
            }
            var aliases = _aliasStore.Load();
            foreach (var kv in aliases) _aliasMap[kv.Key] = kv.Value;

            // Top buttons
            _btnSettings = new Button { Text = "Settings", Size = new Size(92, 32) };
            _btnSettings.Click += SettingsButton_Click;
            Controls.Add(_btnSettings);
            _toolTip.SetToolTip(_btnSettings, "Edit aliases, primary monitor, theme, and app tools.");

            _btnRefresh = new Button { Text = "Refresh", Size = new Size(88, 32) };
            _btnRefresh.Click += (_, __) => RefreshMonitorsAndUi();
            Controls.Add(_btnRefresh);
            _toolTip.SetToolTip(_btnRefresh, "Detect monitors again.");

            _btnSaveLayout = new Button { Text = "Save", Size = new Size(74, 32) };
            _btnSaveLayout.Click += (_, __) => SaveSelectedLayoutProfile();
            Controls.Add(_btnSaveLayout);
            _toolTip.SetToolTip(_btnSaveLayout, "Save the current monitor layout to a named profile.");

            _btnRestoreLayout = new Button { Text = "Restore", Size = new Size(82, 32) };
            _btnRestoreLayout.Click += async (_, __) => await RestoreSelectedLayoutProfileAsync(showMessage: true);
            Controls.Add(_btnRestoreLayout);
            _toolTip.SetToolTip(_btnRestoreLayout, "Restore the selected saved monitor layout.");

            _summaryLabel = new Label
            {
                AutoSize = true,
                Text = "Detecting monitors...",
                Font = new Font("Segoe UI", 8.75f, FontStyle.Regular)
            };
            Controls.Add(_summaryLabel);

            _missingToolPanel = new Panel
            {
                Height = 36,
                Visible = false,
                BorderStyle = BorderStyle.FixedSingle
            };
            _missingToolLabel = new Label
            {
                AutoSize = false,
                Text = "MultiMonitorTool is missing. Detection is limited and monitor actions are unavailable.",
                TextAlign = ContentAlignment.MiddleLeft
            };
            _missingToolSettingsButton = new Button { Text = "Settings", Size = new Size(82, 26) };
            _missingToolSettingsButton.Click += SettingsButton_Click;
            _missingToolPanel.Controls.Add(_missingToolLabel);
            _missingToolPanel.Controls.Add(_missingToolSettingsButton);
            Controls.Add(_missingToolPanel);

            SetupTrayIcon();
            SetupSingleInstanceRestoreSignal();

            _hrTop = new Panel { Height = 1, BackColor = SystemColors.ControlDark, Visible = true };
            Controls.Add(_hrTop);

            SizeChanged += (_, __) => LeftTopButtons();
            Layout += (_, __) => LeftTopButtons();

            CheckToolExists();

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

            FormClosing += (_, e) =>
            {
                if (!_exitRequested && _uiSettings.MinimizeToTray && e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                    HideToTray();
                    return;
                }

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

            if (!_restoringFromTray && _uiSettings != null && _uiSettings.MinimizeToTray && WindowState == FormWindowState.Minimized && Visible)
            {
                BeginInvoke(new Action(HideToTray));
                return;
            }

            if (IsNormalVisible && _deferredLayout)
            {
                _deferredLayout = false;
                RefreshMonitorsAndUi();
                LeftTopButtons();
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _showExistingWindowWait?.Unregister(null);
            _showExistingWindowEvent?.Dispose();
            _restoreRefreshTimer?.Stop();
            _restoreRefreshTimer?.Dispose();
            _trayIcon?.Dispose();
            _trayMenu?.Dispose();
            base.OnFormClosed(e);
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

            if (m.Msg == SingleInstanceMessenger.ShowExistingWindowMessage)
            {
                ShowMainWindowForInteraction();
                return;
            }

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
        // Prefers the configured fallback primary, then the left-most active screen.
        private Screen GetFallbackActiveScreen(string? excludeDevice)
        {
            var configuredFallback = PrimaryMonitorPreference.ResolveConfiguredFallbackDeviceName(
                excludeDevice,
                _detected,
                _aliasMap);
            if (!string.IsNullOrWhiteSpace(configuredFallback))
            {
                var configuredScreen = TryGetScreenForDevice(configuredFallback);
                if (configuredScreen != null)
                    return configuredScreen;
            }

            foreach (var dev in PrimaryMonitorPreference.ResolveAutomaticFallbackDeviceNames(excludeDevice, _detected))
            {
                var scr = TryGetScreenForDevice(dev);
                if (scr != null) return scr;
            }

            var primary = Screen.PrimaryScreen;
            if (primary != null &&
                !string.Equals(primary.DeviceName, excludeDevice, StringComparison.OrdinalIgnoreCase))
                return primary;

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
            if (_btnSettings == null || _btnRefresh == null || _btnSaveLayout == null || _btnRestoreLayout == null) return;

            // Don't lay out while minimized; mark that we owe a layout later.
            if (!IsNormalVisible) { _deferredLayout = true; return; }

            const int spacing = 18;
            int x = SideMargin;

            _btnSettings.Location = new Point(x, TopButtonsY);
            x = _btnSettings.Right + spacing;

            _btnRefresh.Location = new Point(x, TopButtonsY);
            x = _btnRefresh.Right + spacing;

            _btnSaveLayout.Location = new Point(x, TopButtonsY);
            x = _btnSaveLayout.Right + spacing;

            _btnRestoreLayout.Location = new Point(x, TopButtonsY);
            x = _btnRestoreLayout.Right + spacing;

            _btnSettings.BringToFront();
            _btnRefresh.BringToFront();
            _btnSaveLayout.BringToFront();
            _btnRestoreLayout.BringToFront();

            UpdateTopSeparator();
        }

        private int GetCompactClientWidth()
        {
            int contentRight = SideMargin + RowPanelWidth;
            int toolbarRight = (_btnRestoreLayout?.Right ?? _btnSaveLayout?.Right ?? 0);
            return Math.Max(contentRight + SideMargin, toolbarRight + SideMargin);
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

            UpdateMissingToolPanel();
        }

        private void UpdateMissingToolPanel()
        {
            if (_missingToolPanel == null || _missingToolLabel == null || _missingToolSettingsButton == null)
                return;

            bool missing = !File.Exists(_toolPath);
            _missingToolPanel.Visible = missing;
            if (!missing) return;

            int y = GetSeparatorY() + 8;
            _missingToolPanel.Location = new Point(SideMargin, y);
            _missingToolPanel.Width = Math.Max(0, ClientSize.Width - (SideMargin * 2));

            _missingToolSettingsButton.Location = new Point(_missingToolPanel.Width - _missingToolSettingsButton.Width - 6, 4);
            _missingToolLabel.Location = new Point(10, 5);
            _missingToolLabel.Size = new Size(Math.Max(0, _missingToolSettingsButton.Left - 18), 24);
            _missingToolPanel.BringToFront();
        }

        private int GetRowsStartY()
        {
            int sepY = GetSeparatorY();
            int yStart = sepY + (_hrTop?.Height ?? 1) + 12;
            if (_missingToolPanel?.Visible == true)
                yStart = _missingToolPanel.Bottom + 12;
            return Math.Max(FirstRowY, yStart);
        }

        private void SetupTrayIcon()
        {
            _trayMenu = new ContextMenuStrip();
            _trayMenu.Items.Add("Open Monitor Switcher", null, (_, __) => RestoreFromTray());
            _trayMenu.Items.Add("Refresh", null, (_, __) => RefreshMonitorsAndUi());
            _trayMenu.Items.Add("Save Layout", null, (_, __) =>
            {
                ShowMainWindowForInteraction();
                SaveSelectedLayoutProfile();
            });
            _trayMenu.Items.Add("Restore Layout", null, async (_, __) => await RestoreSelectedLayoutProfileAsync(showMessage: false));
            _trayMenu.Items.Add("Settings", null, (sender, e) =>
            {
                ShowMainWindowForInteraction();
                SettingsButton_Click(sender, e);
            });
            _trayMenu.Items.Add(new ToolStripSeparator());
            _trayMenu.Items.Add("Exit", null, (_, __) =>
            {
                _exitRequested = true;
                Close();
            });

            Icon trayIcon;
            try
            {
                trayIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? SystemIcons.Application;
            }
            catch
            {
                trayIcon = SystemIcons.Application;
            }

            _trayIcon = new NotifyIcon
            {
                Text = "Monitor Switcher",
                ContextMenuStrip = _trayMenu,
                Icon = trayIcon,
                Visible = true
            };

            _trayIcon.DoubleClick += (_, __) => RestoreFromTray();
        }

        private void SetupSingleInstanceRestoreSignal()
        {
            try
            {
                _showExistingWindowEvent = new EventWaitHandle(
                    initialState: false,
                    mode: EventResetMode.AutoReset,
                    name: SingleInstanceMessenger.ShowExistingWindowEventName);

                _showExistingWindowWait = ThreadPool.RegisterWaitForSingleObject(
                    _showExistingWindowEvent,
                    (_, __) =>
                    {
                        if (IsDisposed) return;
                        try { BeginInvoke(new Action(ShowMainWindowForInteraction)); }
                        catch { /* ignore shutdown races */ }
                    },
                    state: null,
                    millisecondsTimeOutInterval: -1,
                    executeOnlyOnce: false);
            }
            catch
            {
                // The window-message fallback still handles normal visible-window activation.
            }
        }

        private void HideToTray()
        {
            SaveWindowBounds();
            _uiStore.Save(_uiSettings);
            ShowInTaskbar = false;
            Hide();
            if (_trayIcon != null)
                _trayIcon.ShowBalloonTip(2000, "Monitor Switcher", "Still running in the notification area.", ToolTipIcon.Info);
        }

        private void ShowMainWindowForInteraction()
        {
            if (!Visible || WindowState == FormWindowState.Minimized || !ShowInTaskbar)
                RestoreFromTray();
            else
                Activate();
        }

        private void RestoreFromTray()
        {
            _restoringFromTray = true;
            ShowInTaskbar = true;
            WindowState = FormWindowState.Normal;
            Show();
            Activate();

            QueueRestoreRefresh();
        }

        private void QueueRestoreRefresh()
        {
            _restoreRefreshTimer?.Stop();
            _restoreRefreshTimer?.Dispose();

            _restoreRefreshTimer = new System.Windows.Forms.Timer { Interval = 350 };
            _restoreRefreshTimer.Tick += (_, __) =>
            {
                _restoreRefreshTimer?.Stop();
                _restoreRefreshTimer?.Dispose();
                _restoreRefreshTimer = null;

                CompleteRestoreFromTray();
            };
            _restoreRefreshTimer.Start();
        }

        private void CompleteRestoreFromTray()
        {
            try
            {
                BeginInvoke(new Action(() =>
                {
                    try
                    {
                        WindowState = FormWindowState.Normal;
                        _deferredLayout = false;
                        RefreshMonitorsAndUi();
                        LeftTopButtons();
                        EnsureDynamicRowsVisible();
                        Invalidate(true);
                        Update();
                    }
                    finally
                    {
                        _restoringFromTray = false;
                    }
                }));
            }
            catch
            {
                _restoringFromTray = false;
            }
        }

        private string SelectedLayoutProfileName()
            => LayoutProfileStore.NormalizeProfileName(_uiSettings.SelectedLayoutProfile);

        private string SelectedLayoutPath()
            => _profileStore.GetLayoutPath(SelectedLayoutProfileName());

        private void SaveSelectedLayoutProfile()
        {
            if (_displayActionGate.CurrentCount == 0)
            {
                ThemedMessageBox.Info(this,
                    "Another monitor action is already running. Wait for it to finish, then save the layout again.",
                    "Monitor Switcher", _uiSettings.DarkMode);
                return;
            }

            var profile = PromptForLayoutProfileName(SelectedLayoutProfileName());
            if (string.IsNullOrWhiteSpace(profile))
                return;

            if (!TryBeginDisplayAction("save layout"))
                return;

            try
            {
                var path = _profileStore.GetLayoutPath(profile);
                var ok = _layoutSvc.SaveLayout(path);
                _log.Write(ok
                    ? $"Saved layout profile '{profile}' to {path}."
                    : $"Failed to save layout profile '{profile}'. MultiMonitorTool present: {File.Exists(_toolPath)}.");

                if (!ok)
                {
                    ThemedMessageBox.Info(this,
                        "Unable to save layout. Is MultiMonitorTool.exe present?",
                        "Save Layout", _uiSettings.DarkMode);
                    return;
                }

                _profileStore.AddProfileName(profile);
                _uiSettings.SelectedLayoutProfile = profile;
                _uiStore.Save(_uiSettings);
                ThemedMessageBox.Info(this, $"Layout profile '{profile}' saved.", "Save Layout", _uiSettings.DarkMode);
            }
            finally
            {
                EndDisplayAction("save layout");
            }
        }

        private string? PromptForLayoutProfileName(string currentName)
        {
            using var dlg = new Form
            {
                Text = "Save Layout Profile",
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MinimizeBox = false,
                MaximizeBox = false,
                ShowInTaskbar = false,
                ClientSize = new Size(360, 126)
            };

            var label = new Label
            {
                Text = "Profile name",
                Location = new Point(14, 16),
                AutoSize = true
            };
            var input = new TextBox
            {
                Text = currentName,
                Location = new Point(14, 42),
                Size = new Size(330, 24)
            };
            var ok = new Button { Text = "Save", DialogResult = DialogResult.OK, Size = new Size(82, 30), Location = new Point(172, 84) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Size = new Size(82, 30), Location = new Point(262, 84) };

            dlg.Controls.Add(label);
            dlg.Controls.Add(input);
            dlg.Controls.Add(ok);
            dlg.Controls.Add(cancel);
            dlg.AcceptButton = ok;
            dlg.CancelButton = cancel;
            Themer.Apply(dlg, _uiSettings.DarkMode ? ThemePalette.Dark() : ThemePalette.Light());
            DwmInterop.SetDarkTitleBar(dlg.Handle, _uiSettings.DarkMode);

            input.SelectAll();
            return dlg.ShowDialog(this) == DialogResult.OK
                ? LayoutProfileStore.NormalizeProfileName(input.Text)
                : null;
        }

        private async Task RestoreSelectedLayoutProfileAsync(bool showMessage)
        {
            if (!TryBeginDisplayAction("restore layout"))
                return;

            try
            {
                var profile = SelectedLayoutProfileName();
                var path = _profileStore.GetLayoutPath(profile);
                var ok = _layoutSvc.LoadLayout(path);
                _log.Write(ok
                    ? $"Restored layout profile '{profile}' from {path}."
                    : $"Failed to restore layout profile '{profile}' from {path}.");

                if (!ok)
                {
                    if (showMessage)
                        ThemedMessageBox.Warn(this,
                            $"Unable to restore layout profile '{profile}'. Save it first and confirm MultiMonitorTool.exe is present.",
                            "Restore Layout", _uiSettings.DarkMode);
                    return;
                }

                await Task.Delay(800);
                await ApplySavedLayoutTopologyWithRetryAsync(profile, path);
                RefreshMonitorsAndUi();
                if (EnforcePreferredPrimaryIfActive("preferred primary after layout restore"))
                {
                    await Task.Delay(400);
                    RefreshMonitorsAndUi();
                }

                if (showMessage)
                    ThemedMessageBox.Info(this, $"Layout profile '{profile}' restored.", "Restore Layout", _uiSettings.DarkMode);
            }
            finally
            {
                EndDisplayAction("restore layout");
            }
        }

        private async Task<DisplayTopologyResult> ApplySavedLayoutTopologyWithRetryAsync(string profile, string path)
        {
            DisplayTopologyResult? result = null;

            for (int attempt = 0; attempt < 5; attempt++)
            {
                result = _topologySvc.ApplyLayoutPositionsFromConfig(path);
                if (result.Success)
                    break;

                await Task.Delay(500);
            }

            result ??= new DisplayTopologyResult
            {
                Success = false,
                Message = "Saved layout topology was not attempted."
            };

            LogTopologyResult($"CCD layout restore for profile '{profile}'", result);
            if (result.Success)
                await Task.Delay(400);

            return result;
        }

        private void LogTopologyResult(string action, DisplayTopologyResult result)
        {
            _log.Write(
                $"{action}: success={result.Success}, validate={FormatCode(result.ValidateCode)}, " +
                $"apply={FormatCode(result.ApplyCode)}, message='{result.Message}'.");

            foreach (var detail in result.Details)
                _log.Write($"{action}: {detail}");
        }

        private static string FormatCode(int? code)
            => code.HasValue ? code.Value.ToString() : "n/a";

        // ---------- Theme / caption ----------

        private void ApplyTheme(bool dark)
        {
            var palette = dark ? ThemePalette.Dark() : ThemePalette.Light();

            // Apply to this form + children
            Themer.Apply(this, palette);

            // Keep HR visible
            if (_hrTop != null) _hrTop.BackColor = palette.Border;
            if (_summaryLabel != null) _summaryLabel.ForeColor = palette.TextSubtle;

            // Native title bar
            DwmInterop.SetDarkTitleBar(this.Handle, dark);
            var accent = WindowsTheme.AccentColor();
            if (accent.HasValue) DwmInterop.SetCaptionColor(this.Handle, accent.Value);

            Invalidate(true);
            Update();
        }

        // ---------- Settings dialog ----------
        private async void SettingsButton_Click(object? sender, EventArgs e)
        {
            if (_displayActionGate.CurrentCount == 0)
            {
                ThemedMessageBox.Info(this,
                    "Another monitor action is already running. Wait for it to finish, then open Settings again.",
                    "Monitor Switcher", _uiSettings.DarkMode);
                return;
            }

            if (!Visible || WindowState == FormWindowState.Minimized || !ShowInTaskbar)
                ShowMainWindowForInteraction();

            using var dlg = new AliasSettingsForm(
                BuildAliasSettingsRows(),
                _uiSettings.DarkMode,
                _uiSettings.AlwaysOnTop,
                _uiSettings.MinimizeToTray,
                _uiSettings.StartWithWindows,
                _uiSettings.ConfirmBeforeDisable,
                Path.Combine(AppContext.BaseDirectory, "MultiMonitorTool.cfg"),
                _profileStore.LoadProfileNames(),
                SelectedLayoutProfileName(),
                _log.Read(),
                this)
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
                try
                {
                    await ApplySettingsDialogResultsAsync(dlg);
                }
                catch (Exception ex)
                {
                    ThemedMessageBox.Error(this,
                        $"Unable to apply settings.\n{ex.Message}",
                        "Monitor Switcher",
                        _uiSettings.DarkMode);
                }
            }
        }

        private List<AliasViewRow> BuildAliasSettingsRows()
            => AliasSettingsMapper.BuildRows(BuildPresentationList(), _aliasMap);

        private async Task ApplySettingsDialogResultsAsync(AliasSettingsForm dlg)
        {
            if (!TryBeginDisplayAction("apply settings"))
                return;

            try
            {
                AliasSettingsMapper.ApplyMonitorSettings(
                    _aliasMap,
                    dlg.RemovedKeys,
                    dlg.UpdatedMappings,
                    dlg.PreferredPrimaryKey,
                    dlg.FallbackPrimaryKey);
                _aliasStore.Save(_aliasMap);

                if (dlg.DarkModeResult.HasValue && dlg.DarkModeResult.Value != _uiSettings.DarkMode)
                {
                    _uiSettings.DarkMode = dlg.DarkModeResult.Value;
                    _uiStore.Save(_uiSettings);
                    ApplyTheme(_uiSettings.DarkMode);
                }

                if (dlg.AlwaysOnTopResult.HasValue && dlg.AlwaysOnTopResult.Value != _uiSettings.AlwaysOnTop)
                {
                    _uiSettings.AlwaysOnTop = dlg.AlwaysOnTopResult.Value;
                    _uiStore.Save(_uiSettings);
                    TopMost = _uiSettings.AlwaysOnTop;
                }

                if (dlg.MinimizeToTrayResult.HasValue)
                    _uiSettings.MinimizeToTray = dlg.MinimizeToTrayResult.Value;

                if (dlg.StartWithWindowsResult.HasValue)
                {
                    _uiSettings.StartWithWindows = dlg.StartWithWindowsResult.Value;
                    StartupManager.SetEnabled(_uiSettings.StartWithWindows, Application.ExecutablePath);
                }

                if (dlg.ConfirmBeforeDisableResult.HasValue)
                    _uiSettings.ConfirmBeforeDisable = dlg.ConfirmBeforeDisableResult.Value;

                if (!string.IsNullOrWhiteSpace(dlg.SelectedLayoutProfileResult))
                    _uiSettings.SelectedLayoutProfile = dlg.SelectedLayoutProfileResult;

                foreach (var profile in dlg.RemovedLayoutProfiles)
                    _profileStore.DeleteProfile(profile);

                _uiStore.Save(_uiSettings);

                RefreshMonitorsAndUi();
                if (EnforcePreferredPrimaryIfActive("preferred primary after settings save"))
                {
                    await Task.Delay(400);
                    RefreshMonitorsAndUi();
                }
            }
            finally
            {
                EndDisplayAction("apply settings");
            }
        }

        private bool TryBeginDisplayAction(string actionName)
        {
            if (_displayActionGate.Wait(0))
            {
                _log.Write($"Display action started: {actionName}.");
                return true;
            }

            ThemedMessageBox.Info(this,
                "Another monitor action is already running. Wait for it to finish, then try again.",
                "Monitor Switcher", _uiSettings.DarkMode);
            return false;
        }

        private void EndDisplayAction(string actionName)
        {
            _log.Write($"Display action finished: {actionName}.");
            _displayActionGate.Release();
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
            if (ctrls.StatusLabel.AutoSize)
            {
                var statusSize = ctrls.StatusLabel.PreferredSize;
                int statusX = ctrls.EnableButton.Right - statusSize.Width;
                int statusY = ctrls.TitleLabel.Top + Math.Max(0, (ctrls.TitleLabel.Height - statusSize.Height) / 2);
                ctrls.StatusLabel.Location = new Point(statusX, statusY);
            }

            ctrls.StatusLabel.Invalidate();
        }

        // ---------- Refresh + dynamic UI build ----------

        private void RefreshMonitorsAndUi()
        {
            if (_isRefreshingUi)
            {
                _refreshAgainAfterCurrent = true;
                return;
            }

            _isRefreshingUi = true;
            try
            {
                do
                {
                    _refreshAgainAfterCurrent = false;
                    RefreshMonitorsAndUiCore();
                }
                while (_refreshAgainAfterCurrent && !IsDisposed);
            }
            finally
            {
                _isRefreshingUi = false;
            }
        }

        private void RefreshMonitorsAndUiCore()
        {
            LeftTopButtons();
            UpdateTopSeparator();

            if (!IsNormalVisible)
            {
                _deferredLayout = true;
                return;
            }

            _detected = _detectSvc.Detect();

            // Attempt to re-bind aliases if stable keys changed (e.g., port swaps)
            ReconcileAliasesForDetected();
            SuppressShadowedDeviceFallbackDetections();
            ReconcilePositionalAliases();
            LogDetectionSnapshotIfChanged();

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
                if (m.IsPresent && m.IsActive && !string.IsNullOrWhiteSpace(m.DeviceName))
                    info.LastDeviceName = m.DeviceName;
                if (m.IsPresent && m.IsActive)
                    info.LastKnownX = m.PositionX;
                if (!string.IsNullOrWhiteSpace(m.MonitorKey))
                    info.LastRegistryKey = m.MonitorKey;
                if (!string.IsNullOrWhiteSpace(m.SerialNumber))
                    info.LastSerialNumber = m.SerialNumber;
                if (!string.IsNullOrWhiteSpace(m.InstanceId))
                    info.LastInstanceId = m.InstanceId;
                if (!string.IsNullOrWhiteSpace(m.MonitorId))
                    info.LastMonitorId = m.MonitorId;

                if (m.IsPresent && m.IsActive)
                    MonitorTargetResolver.EnsureKnownTargets(_aliasMap, m.StableKey, m.DeviceName, m.Name);
            }

            RemoveShadowedDeviceAliases();
            MonitorTargetResolver.PruneKnownTargets(_aliasMap, _detected);
            _aliasStore.Save(_aliasMap);

            var toShow = BuildPresentationList();
            UpdateMonitorSummary(toShow);

            SuspendLayout();
            try
            {
                foreach (var c in _dynamicControls) { Controls.Remove(c); c.Dispose(); }
                _dynamicControls.Clear();
                _controlsByKey.Clear();

                _layoutRightMost = 0;

                UpdateMissingToolPanel();
                int y = GetRowsStartY();

                int lastRowBottom = y;
                foreach (var m in toShow)
                {
                    AddMonitorControls(m, GetAliasFor(m.StableKey), y);
                    lastRowBottom = y + RowPanelHeight;
                    y += RowVerticalGap;
                }

                if (_summaryLabel != null)
                {
                    var summarySize = _summaryLabel.PreferredSize;
                    int summaryX = SideMargin + RowPanelWidth - summarySize.Width;
                    int summaryY = Math.Max(GetRowsStartY(), lastRowBottom + 4);
                    _summaryLabel.Location = new Point(summaryX, summaryY);
                    _summaryLabel.BringToFront();
                    y = _summaryLabel.Bottom + 8;
                }

                int desiredWidth = GetCompactClientWidth();
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
                if (_summaryLabel != null) _summaryLabel.ForeColor = palette.TextSubtle;


                UpdateButtonStatus();
                EnsureDynamicRowsVisible();
            }
            finally
            {
                ResumeLayout(performLayout: true);
            }
        }

        private void EnsureDynamicRowsVisible()
        {
            if (_dynamicControls.Count == 0)
                return;

            foreach (var control in _dynamicControls.Where(c => !c.IsDisposed))
            {
                control.Visible = true;
                EnsureControlTreeCreated(control);
                control.BringToFront();
                control.Invalidate(true);
                ForceRedraw(control);
            }

            _summaryLabel?.BringToFront();
            _btnSettings?.BringToFront();
            _btnRefresh?.BringToFront();
            _btnSaveLayout?.BringToFront();
            _btnRestoreLayout?.BringToFront();
            _hrTop?.BringToFront();

            ForceRedraw(this);
        }

        private static void EnsureControlTreeCreated(Control control)
        {
            control.CreateControl();
            foreach (Control child in control.Controls)
                EnsureControlTreeCreated(child);
        }

        private static void ForceRedraw(Control control)
        {
            if (!control.IsHandleCreated) return;

            const int RDW_INVALIDATE = 0x0001;
            const int RDW_ERASE = 0x0004;
            const int RDW_ALLCHILDREN = 0x0080;
            const int RDW_UPDATENOW = 0x0100;

            RedrawWindow(
                control.Handle,
                IntPtr.Zero,
                IntPtr.Zero,
                RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        }

        [DllImport("user32.dll")]
        private static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, int flags);

        private void UpdateMonitorSummary(IReadOnlyCollection<DetectedMonitor> monitors)
        {
            if (_summaryLabel == null) return;

            int present = monitors.Count(m => m.IsPresent);
            int active = monitors.Count(m => m.IsPresent && m.IsActive);
            int savedOnly = monitors.Count - present;

            _summaryLabel.Text = savedOnly > 0
                ? $"{active}/{present} active - {savedOnly} saved"
                : $"{active}/{present} active";
        }

        private List<DetectedMonitor> BuildPresentationList()
            => MonitorPresentationBuilder.Build(_detected, _aliasMap);

        private void AddMonitorControls(DetectedMonitor monitor, string friendlyName, int positionY)
        {
            var palette = _uiSettings.DarkMode ? ThemePalette.Dark() : ThemePalette.Light();

            var card = new Panel
            {
                Location = new Point(SideMargin, positionY),
                Size = new Size(RowPanelWidth, RowPanelHeight),
                BackColor = palette.Surface,
                BorderStyle = BorderStyle.FixedSingle
            };
            Controls.Add(card);
            _dynamicControls.Add(card);

            int labelX = 14;
            int disableX = 14;
            int enableX = disableX + ButtonWidth + ControlGapX;
            int buttonY = 42;

            var label = new Label
            {
                Text = friendlyName,
                Location = new Point(labelX, 12),
                AutoSize = false,
                Size = new Size(245, 22),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                AutoEllipsis = true
            };
            card.Controls.Add(label);
            _toolTip.SetToolTip(label, friendlyName);

            var statusLabel = new Label
            {
                AutoSize = false,
                Location = new Point(RowPanelWidth - 103, 12),
                Size = new Size(86, 22),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 8.25f, FontStyle.Bold),
                BorderStyle = BorderStyle.FixedSingle
            };
            card.Controls.Add(statusLabel);

            var detail = new Label
            {
                Text = BuildMonitorDetailText(monitor),
                Location = new Point(246, 46),
                AutoSize = false,
                Size = new Size(126, 18),
                AutoEllipsis = true,
                Font = new Font("Segoe UI", 8.25f, FontStyle.Regular),
                ForeColor = palette.TextSubtle,
                TextAlign = ContentAlignment.MiddleRight
            };
            card.Controls.Add(detail);
            _toolTip.SetToolTip(detail, BuildMonitorTooltipText(monitor));

            var buttonOff = new Button
            {
                Text = "Disable",
                Location = new Point(disableX, buttonY),
                Size = new Size(ButtonWidth, ButtonHeight),
                Tag = monitor.StableKey
            };
            buttonOff.Click += ButtonOff_Click;
            card.Controls.Add(buttonOff);
            _toolTip.SetToolTip(buttonOff, "Disable this monitor. At least one display must remain active.");

            var buttonOn = new Button
            {
                Text = "Enable",
                Location = new Point(enableX, buttonY),
                Size = new Size(ButtonWidth, ButtonHeight),
                Tag = monitor.StableKey
            };
            buttonOn.Click += ButtonOn_Click;
            card.Controls.Add(buttonOn);
            _toolTip.SetToolTip(buttonOn, "Enable this monitor using the best known monitor identifier.");

            statusLabel.Text = monitor.IsPresent ? (monitor.IsActive ? "ONLINE" : "DISABLED") : "OFFLINE";
            statusLabel.ForeColor = monitor.IsPresent ? (monitor.IsActive ? palette.StatusOk : palette.StatusWarn) : palette.TextSubtle;
            statusLabel.BackColor = palette.Back;

            _layoutRightMost = Math.Max(_layoutRightMost, card.Right);

            _controlsByKey[monitor.StableKey] = new MonitorControls
            {
                DisableButton = buttonOff,
                EnableButton = buttonOn,
                StatusLabel = statusLabel,
                TitleLabel = label
            };
        }

        private static string BuildMonitorDetailText(DetectedMonitor monitor)
        {
            if (!string.IsNullOrWhiteSpace(monitor.DeviceName))
                return monitor.DeviceName.Replace(@"\\.\", string.Empty);
            if (!string.IsNullOrWhiteSpace(monitor.Name))
                return monitor.Name;
            return monitor.IsPresent ? "Detected" : "Saved";
        }

        private static string BuildMonitorTooltipText(DetectedMonitor monitor)
        {
            var lines = new List<string>();
            if (!string.IsNullOrWhiteSpace(monitor.DeviceName)) lines.Add($"Device: {monitor.DeviceName}");
            if (!string.IsNullOrWhiteSpace(monitor.Name)) lines.Add($"Name: {monitor.Name}");
            if (!string.IsNullOrWhiteSpace(monitor.MonitorKey)) lines.Add($"Registry: {monitor.MonitorKey}");
            lines.Add($"Stable key: {monitor.StableKey}");
            return string.Join(Environment.NewLine, lines);
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

            if (_uiSettings.ConfirmBeforeDisable)
            {
                var name = GetAliasFor(stableKey);
                var choice = MessageBox.Show(
                    this,
                    $"Disable '{name}'?",
                    "Disable Monitor",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (choice != DialogResult.Yes)
                    return;
            }

            if (!TryBeginDisplayAction("disable monitor"))
                return;

            try
            {
                var target = MonitorTargetResolver.ResolveDisableTargetArg(stableKey, _detected, _aliasMap);
                if (string.IsNullOrWhiteSpace(target))
                {
                    LogMonitorAction($"DISABLE {stableKey}: no target resolved.");
                    ThemedMessageBox.Info(this,
                        "Cannot disable: no resolvable device identifier for this monitor.",
                        "Monitor Switcher", _uiSettings.DarkMode);
                    return;
                }
                LogMonitorAction($"DISABLE {stableKey}: selected target '{target}'.");
                _log.Write($"Disable requested for '{GetAliasFor(stableKey)}' using target '{target}'.");

                var disablingDevice = TryGetDeviceNameForStableKey(stableKey);

                // Show WORKING… and lock the row
                SetRowBusy(stableKey, true);
                UpdateButtonStatus();

                bool disabledByTopology = false;
                bool disabledByTool = false;
                bool failedBeforeDisable = false;

                try
                {
                    // Capture baseline layout before any primary-monitor topology change.
                    bool allActiveNow = _detected.Count > 0 && _detected.All(d => d.IsPresent && d.IsActive);
                    if (allActiveNow) _layoutSvc.SaveLayout(SelectedLayoutPath());

                    if (!string.IsNullOrWhiteSpace(disablingDevice))
                    {
                        MoveWindowIfHostedOn(disablingDevice);
                        await Task.Delay(100);

                        var primaryDisable = await DisablePrimaryUsingTopologyIfNeededAsync(disablingDevice);
                        if (primaryDisable.Attempted)
                        {
                            if (!primaryDisable.Success)
                            {
                                failedBeforeDisable = true;
                                ThemedMessageBox.Error(this,
                                    primaryDisable.ErrorMessage,
                                    "Monitor Switcher", _uiSettings.DarkMode);
                            }
                            else
                            {
                                disabledByTopology = true;
                            }
                        }
                    }

                    if (!failedBeforeDisable && !disabledByTopology)
                    {
                        var res = await ExecToolAsync("/disable", target);
                        LogMonitorAction($"DISABLE {stableKey}: exit={res.ExitCode}, stderr='{res.StdErr}'.");
                        _log.Write($"Disable result for '{GetAliasFor(stableKey)}': exit={res.ExitCode}, stderr='{res.StdErr}'.");

                        var detectedAfterTool = await WaitForDisableDetectionAsync(stableKey);
                        disabledByTool = !MonitorTargetResolver.IsStableKeyActive(detectedAfterTool, stableKey);

                        if (!disabledByTool && !string.IsNullOrWhiteSpace(disablingDevice))
                        {
                            _log.Write(
                                $"Disable for '{GetAliasFor(stableKey)}' left the monitor active after MultiMonitorTool exit {res.ExitCode}; trying CCD topology fallback.");

                            var fallbackDisable = await DisableUsingTopologyAsync(
                                disablingDevice,
                                "CCD disable fallback after MultiMonitorTool");

                            disabledByTopology = fallbackDisable.Success;
                            if (!fallbackDisable.Success)
                            {
                                ThemedMessageBox.Error(this,
                                    $"Disable failed. MultiMonitorTool exit {res.ExitCode}; {fallbackDisable.ErrorMessage}" +
                                    (string.IsNullOrWhiteSpace(res.StdErr) ? string.Empty : $"\n{res.StdErr}"),
                                    "Monitor Switcher", _uiSettings.DarkMode);
                            }
                        }
                        else if (!disabledByTool)
                        {
                            ThemedMessageBox.Error(this,
                                $"Disable failed (exit {res.ExitCode})." +
                                (string.IsNullOrWhiteSpace(res.StdErr) ? string.Empty : $"\n{res.StdErr}"),
                                "Monitor Switcher", _uiSettings.DarkMode);
                        }
                    }
                    else if (disabledByTopology)
                    {
                        LogMonitorAction($"DISABLE {stableKey}: completed by CCD topology update.");
                        _log.Write($"Disable result for '{GetAliasFor(stableKey)}': completed by CCD topology update.");
                    }

                    if (disabledByTool)
                        _log.Write($"Disable result for '{GetAliasFor(stableKey)}': monitor inactive after MultiMonitorTool command.");

                    // Give the desktop a brief moment to settle.
                    await Task.Delay(400);
                }
                finally
                {
                    SetRowBusy(stableKey, false);
                }

                RefreshMonitorsAndUi();
            }
            finally
            {
                EndDisplayAction("disable monitor");
            }
        }

        private async Task<PrimaryDisableAttempt> DisablePrimaryUsingTopologyIfNeededAsync(string disablingDevice)
        {
            var primary = Screen.PrimaryScreen;
            if (primary == null ||
                !primary.DeviceName.Equals(disablingDevice, StringComparison.OrdinalIgnoreCase))
                return new PrimaryDisableAttempt { Attempted = false, Success = true };

            var result = await DisableUsingTopologyAsync(disablingDevice, $"CCD primary disable {disablingDevice}");

            return new PrimaryDisableAttempt
            {
                Attempted = true,
                Success = result.Success,
                ErrorMessage = result.ErrorMessage
            };
        }

        private async Task<PrimaryDisableAttempt> DisableUsingTopologyAsync(string disablingDevice, string logAction)
        {
            var fallback = GetTopologyFallbackScreenForDisable(disablingDevice);
            if (fallback.DeviceName.Equals(disablingDevice, StringComparison.OrdinalIgnoreCase))
            {
                return new PrimaryDisableAttempt
                {
                    Attempted = true,
                    Success = false,
                    ErrorMessage = "Cannot disable the monitor because no fallback monitor is available."
                };
            }

            _log.Write($"Disabling monitor {disablingDevice} with fallback primary {fallback.DeviceName}.");
            var result = _topologySvc.DisableDisplayUsingFallbackPrimary(disablingDevice, fallback.DeviceName);
            LogTopologyResult(logAction, result);
            await Task.Delay(700);

            return new PrimaryDisableAttempt
            {
                Attempted = true,
                Success = result.Success,
                ErrorMessage = result.Success
                    ? string.Empty
                    : $"Disable failed before the monitor could be deactivated. {result.Message}"
            };
        }

        private Screen GetTopologyFallbackScreenForDisable(string disablingDevice)
        {
            var primary = Screen.PrimaryScreen;
            if (primary != null &&
                !primary.DeviceName.Equals(disablingDevice, StringComparison.OrdinalIgnoreCase))
            {
                return primary;
            }

            return GetFallbackActiveScreen(disablingDevice);
        }


        private async void ButtonOn_Click(object? sender, EventArgs e)
        {
            if (sender is not Button btn || btn.Tag is not string stableKey) return;

            if (!TryBeginDisplayAction("enable monitor"))
                return;

            try
            {
                var enableTargets = MonitorTargetResolver.ResolveEnableTargetArgs(stableKey, _detected, _aliasMap).ToList();
                if (enableTargets.Count == 0)
                {
                    LogMonitorAction($"ENABLE {stableKey}: no targets resolved.");
                    ThemedMessageBox.Info(this,
                        "Cannot enable: no resolvable monitor identifiers for this monitor. Toggle another display on, then press Refresh.",
                        "Monitor Switcher", _uiSettings.DarkMode);
                    return;
                }
                LogMonitorAction($"ENABLE {stableKey}: candidate targets [{string.Join(", ", enableTargets.Select(t => $"'{t}'"))}].");
                _log.Write($"Enable requested for '{GetAliasFor(stableKey)}' with {enableTargets.Count} candidate target(s).");

                // Show WORKING… and lock the row
                SetRowBusy(stableKey, true);
                UpdateButtonStatus();

                EnableAttemptResult attempt;
                try
                {
                    attempt = await TryEnableMonitorTargetsAsync(stableKey, enableTargets);
                }
                finally
                {
                    SetRowBusy(stableKey, false);
                }

                if (!attempt.Success)
                {
                    LogMonitorAction($"ENABLE {stableKey}: all target attempts failed.");
                    ThemedMessageBox.Error(this,
                        $"Enable failed to activate the selected monitor." +
                        (attempt.LastResult == null ? string.Empty : $" (last exit {attempt.LastResult.ExitCode}).") +
                        (string.IsNullOrWhiteSpace(attempt.LastResult?.StdErr) ? string.Empty : $"\n{attempt.LastResult!.StdErr}"),
                        "Monitor Switcher", _uiSettings.DarkMode);

                    RefreshMonitorsAndUi();
                    return;
                }

                await FinaliseSuccessfulEnableAsync();
            }
            finally
            {
                EndDisplayAction("enable monitor");
            }
        }

        private async Task<EnableAttemptResult> TryEnableMonitorTargetsAsync(string stableKey, List<string> enableTargets)
        {
            ToolResult? last = null;

            for (int i = 0; i < enableTargets.Count; i++)
            {
                var target = enableTargets[i];
                LogMonitorAction($"ENABLE {stableKey}: trying target '{target}'.");
                last = await ExecToolAsync("/enable", target);
                LogMonitorAction($"ENABLE {stableKey}: target '{target}' exit={last.ExitCode}, stderr='{last.StdErr}'.");
                _log.Write($"Enable attempt for '{GetAliasFor(stableKey)}' target '{target}': exit={last.ExitCode}, stderr='{last.StdErr}'.");

                var detectedNow = await WaitForEnableDetectionAsync(stableKey);
                if (MonitorTargetResolver.IsStableKeyActive(detectedNow, stableKey))
                {
                    LogMonitorAction($"ENABLE {stableKey}: activation confirmed after target '{target}'.");
                    return new EnableAttemptResult { Success = true, LastResult = last };
                }

                AddEnableCandidatesFromDetection(stableKey, enableTargets, detectedNow);
                LogMonitorAction($"ENABLE {stableKey}: target '{target}' did not activate requested stable key.");
            }

            return new EnableAttemptResult { Success = false, LastResult = last };
        }

        private void AddEnableCandidatesFromDetection(
            string stableKey,
            List<string> enableTargets,
            IReadOnlyCollection<DetectedMonitor> detectedNow)
        {
            foreach (var nextTarget in MonitorTargetResolver.ResolveEnableTargetArgs(stableKey, detectedNow, _aliasMap))
            {
                if (enableTargets.Any(existing => MonitorTargetResolver.TargetsEquivalent(existing, nextTarget)))
                    continue;

                enableTargets.Add(nextTarget);
                _log.Write($"Enable candidate added for '{GetAliasFor(stableKey)}' after re-detect: '{nextTarget}'.");
            }
        }

        private async Task FinaliseSuccessfulEnableAsync()
        {
            RefreshMonitorsAndUi();
            bool restoredLayout = false;
            if (_detected.Count > 0 && _detected.All(d => d.IsPresent && d.IsActive))
            {
                var profile = SelectedLayoutProfileName();
                var path = SelectedLayoutPath();
                if (!_layoutSvc.LoadLayout(path))
                {
                    // Optional: notify if missing
                    // ThemedMessageBox.Warn(this, "Saved monitor layout file not found.", "Monitor Switcher", _uiSettings.DarkMode);
                }
                else
                {
                    _log.Write($"Auto-restored layout profile '{profile}' after all monitors became active.");
                    await Task.Delay(800);
                    await ApplySavedLayoutTopologyWithRetryAsync(profile, path);
                    restoredLayout = true;
                }

                RefreshMonitorsAndUi();
            }

            if (EnforcePreferredPrimaryIfActive("preferred primary after enable"))
            {
                await Task.Delay(400);
                RefreshMonitorsAndUi();
            }
            else if (!restoredLayout)
            {
                EnforcePrimaryMonitorOrder(allowAutomaticFallback: true);
            }
        }

        private async Task<List<DetectedMonitor>> WaitForEnableDetectionAsync(string stableKey)
        {
            var latest = new List<DetectedMonitor>();

            for (int attempt = 0; attempt < 8; attempt++)
            {
                await Task.Delay(500);
                latest = _detectSvc.Detect();

                if (MonitorTargetResolver.IsStableKeyActive(latest, stableKey))
                    break;

                if (latest.Any(d =>
                        d.StableKey.Equals(stableKey, StringComparison.OrdinalIgnoreCase) &&
                        d.IsPresent &&
                        !d.IsActive &&
                        !string.IsNullOrWhiteSpace(d.DeviceName)))
                {
                    break;
                }
            }

            return latest;
        }

        private async Task<List<DetectedMonitor>> WaitForDisableDetectionAsync(string stableKey)
        {
            var latest = new List<DetectedMonitor>();

            for (int attempt = 0; attempt < 6; attempt++)
            {
                await Task.Delay(500);
                latest = _detectSvc.Detect();

                if (!MonitorTargetResolver.IsStableKeyActive(latest, stableKey))
                    break;
            }

            return latest;
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

                    if (ctrls.StatusLabel.AutoSize)
                    {
                        var sizeBusy = ctrls.StatusLabel.PreferredSize;
                        int sxBusy = ctrls.EnableButton.Right - sizeBusy.Width;
                        int syBusy = ctrls.TitleLabel.Top + Math.Max(0, (ctrls.TitleLabel.Height - sizeBusy.Height) / 2);
                        ctrls.StatusLabel.Location = new Point(sxBusy, syBusy);
                    }

                    continue;
                }

                var live = present.FirstOrDefault(d => d.StableKey.Equals(key, StringComparison.OrdinalIgnoreCase));
                var info = _aliasMap.TryGetValue(key, out var mi) ? mi : null;

                bool isPresent = live != null;
                bool isActive = live?.IsActive == true;
                ctrls.StatusLabel.Text = isPresent ? (isActive ? "ONLINE" : "DISABLED") : "OFFLINE";
                ctrls.StatusLabel.ForeColor = isPresent
                    ? (isActive ? palette.StatusOk : palette.StatusWarn)
                    : palette.TextSubtle;

                if (ctrls.StatusLabel.AutoSize)
                {
                    var statusSize = ctrls.StatusLabel.PreferredSize;
                    int statusX = ctrls.EnableButton.Right - statusSize.Width;
                    int statusY = ctrls.TitleLabel.Top + Math.Max(0, (ctrls.TitleLabel.Height - statusSize.Height) / 2);
                    ctrls.StatusLabel.Location = new Point(statusX, statusY);
                }

                bool canDisable = isPresent && isActive && activeCount > 1;
                bool canEnable = !isActive && MonitorTargetResolver.ResolveEnableTargetArgs(key, _detected, _aliasMap).Any();

                ctrls.DisableButton.Enabled = canDisable;
                ctrls.DisableButton.ForeColor = canDisable ? ForeColor : palette.TextSubtle;

                ctrls.EnableButton.Enabled = canEnable;
                ctrls.EnableButton.ForeColor = canEnable ? ForeColor : palette.TextSubtle;
            }
        }


        private void EnforcePrimaryMonitorOrder(bool allowAutomaticFallback)
        {
            if (EnforcePreferredPrimaryIfActive("preferred primary"))
                return;

            if (!allowAutomaticFallback)
                return;

            var target = PrimaryMonitorPreference.ResolveLeftMostActiveTarget(_detected, _aliasMap);
            if (!string.IsNullOrWhiteSpace(target))
                SetPrimaryWithTopologyFallback(target, "left-most active primary");
        }

        private bool EnforcePreferredPrimaryIfActive(string reason)
        {
            var target = PrimaryMonitorPreference.ResolvePreferredPrimaryTarget(_detected, _aliasMap);
            if (string.IsNullOrWhiteSpace(target))
                return false;

            SetPrimaryWithTopologyFallback(target, reason);
            return true;
        }

        private void SetPrimaryWithTopologyFallback(string target, string reason)
        {
            var result = _topologySvc.SetPrimaryDisplay(target);
            LogTopologyResult($"CCD set primary ({reason})", result);
            if (!result.Success)
                _layoutSvc.SetPrimary(target);
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

        private void RemoveShadowedDeviceAliases()
        {
            if (_aliasMap.Count == 0 || _detected.Count == 0) return;

            var detectedKeys = new HashSet<string>(_detected.Select(d => d.StableKey), StringComparer.OrdinalIgnoreCase);
            var presentByDevice = _detected
                .Where(d => d.IsPresent && !string.IsNullOrWhiteSpace(d.DeviceName))
                .Select(d => new
                {
                    Monitor = d,
                    Device = NormalizeDeviceNameForComparison(d.DeviceName)
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.Device))
                .GroupBy(x => x.Device, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() == 1)
                .ToDictionary(g => g.Key, g => g.First().Monitor, StringComparer.OrdinalIgnoreCase);

            if (presentByDevice.Count == 0) return;

            foreach (var kv in _aliasMap.ToList())
            {
                var staleKey = kv.Key;
                if (detectedKeys.Contains(staleKey) || !IsDeviceFallbackStableKey(staleKey))
                    continue;

                var source = kv.Value;
                var keyDevice = NormalizeDeviceNameForComparison(staleKey);
                var lastDevice = NormalizeDeviceNameForComparison(source.LastDeviceName);

                DetectedMonitor? live = null;
                if (!string.IsNullOrWhiteSpace(keyDevice))
                    presentByDevice.TryGetValue(keyDevice, out live);
                if (live == null && !string.IsNullOrWhiteSpace(lastDevice))
                    presentByDevice.TryGetValue(lastDevice, out live);
                if (live == null || live.StableKey.Equals(staleKey, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!_aliasMap.TryGetValue(live.StableKey, out var target))
                {
                    target = new MonitorInfo();
                    _aliasMap[live.StableKey] = target;
                }

                MergeMonitorInfo(target, source, live.StableKey);
                if (source.IsPreferredPrimary)
                    EnsureOnlyPreferredPrimary(live.StableKey);
                if (source.IsFallbackPrimary)
                    EnsureOnlyFallbackPrimary(live.StableKey);

                _aliasMap.Remove(staleKey);
                _log.Write($"Removed stale saved display key '{staleKey}' because '{live.StableKey}' is present on {live.DeviceName}.");
            }
        }

        private void SuppressShadowedDeviceFallbackDetections()
        {
            if (_aliasMap.Count == 0 || _detected.Count == 0) return;

            var detectedKeys = new HashSet<string>(_detected.Select(d => d.StableKey), StringComparer.OrdinalIgnoreCase);
            var missingSavedHardwareAliases = _aliasMap
                .Where(kv =>
                    !IsDeviceFallbackStableKey(kv.Key) &&
                    !detectedKeys.Contains(kv.Key) &&
                    HasRestoreTarget(kv.Value))
                .Select(kv => kv.Value)
                .ToList();

            if (missingSavedHardwareAliases.Count == 0)
                return;

            var suppressed = _detected
                .Where(d =>
                    IsDeviceFallbackStableKey(d.StableKey) &&
                    !d.IsActive &&
                    IsGeneratedDeviceFallbackAlias(d.StableKey) &&
                    MatchesMissingSavedHardwareAliasTarget(d, missingSavedHardwareAliases))
                .ToList();

            if (suppressed.Count == 0)
                return;

            var suppressedKeys = new HashSet<string>(
                suppressed.Select(d => d.StableKey),
                StringComparer.OrdinalIgnoreCase);

            _detected = _detected
                .Where(d => !suppressedKeys.Contains(d.StableKey))
                .ToList();

            foreach (var key in suppressedKeys)
            {
                if (_aliasMap.TryGetValue(key, out var info) && IsGeneratedDeviceFallbackAlias(key, info))
                    _aliasMap.Remove(key);
            }

            foreach (var monitor in suppressed)
            {
                _log.Write(
                    $"Suppressed anonymous display fallback '{monitor.StableKey}' on {monitor.DeviceName} because a saved hardware monitor is currently missing.");
            }
        }

        private static bool IsDeviceFallbackStableKey(string stableKey)
            => stableKey.Trim().StartsWith("DEV:", StringComparison.OrdinalIgnoreCase);

        private static bool HasRestoreTarget(MonitorInfo info)
            => !string.IsNullOrWhiteSpace(info.LastDeviceName) ||
               info.KnownTargets.Any(t => !string.IsNullOrWhiteSpace(t));

        private static bool MatchesMissingSavedHardwareAliasTarget(DetectedMonitor monitor, IReadOnlyCollection<MonitorInfo> missingAliases)
        {
            var device = NormalizeDeviceNameForComparison(monitor.DeviceName);
            if (string.IsNullOrWhiteSpace(device))
                device = NormalizeDeviceNameForComparison(monitor.StableKey);
            if (string.IsNullOrWhiteSpace(device))
                return false;

            return missingAliases.Any(info =>
                DeviceTargetEquals(info.LastDeviceName, device) ||
                info.KnownTargets.Any(t => IsLikelyDeviceName(t) && DeviceTargetEquals(t, device)));
        }

        private static bool DeviceTargetEquals(string? target, string device)
            => string.Equals(NormalizeDeviceNameForComparison(target), device, StringComparison.OrdinalIgnoreCase);

        private bool IsGeneratedDeviceFallbackAlias(string stableKey)
        {
            return !_aliasMap.TryGetValue(stableKey, out var info) ||
                   IsGeneratedDeviceFallbackAlias(stableKey, info);
        }

        private static bool IsGeneratedDeviceFallbackAlias(string stableKey, MonitorInfo info)
        {
            return IsPlaceholderAlias(info.Name, stableKey) &&
                   string.IsNullOrWhiteSpace(info.LastSerialNumber) &&
                   string.IsNullOrWhiteSpace(info.LastInstanceId) &&
                   string.IsNullOrWhiteSpace(info.LastRegistryKey) &&
                   string.IsNullOrWhiteSpace(info.LastMonitorId);
        }

        private static string NormalizeDeviceNameForComparison(string? value)
        {
            var t = NormaliseTarget(value ?? string.Empty).Trim();
            if (t.StartsWith("DEV:", StringComparison.OrdinalIgnoreCase))
                t = t[4..].Trim();

            t = t.Replace('/', '\\');
            while (t.Contains("\\\\", StringComparison.Ordinal))
                t = t.Replace("\\\\", "\\");
            return t;
        }

        private static void MergeMonitorInfo(MonitorInfo target, MonitorInfo source, string targetStableKey)
        {
            if (IsPlaceholderAlias(target.Name, targetStableKey) && !string.IsNullOrWhiteSpace(source.Name))
                target.Name = source.Name;

            if (source.IsPreferredPrimary)
                target.IsPreferredPrimary = true;
            if (source.IsFallbackPrimary)
                target.IsFallbackPrimary = true;

            if (!target.PreferredOrder.HasValue && source.PreferredOrder.HasValue)
                target.PreferredOrder = source.PreferredOrder;
            if (!target.LastKnownX.HasValue && source.LastKnownX.HasValue)
                target.LastKnownX = source.LastKnownX;

            target.LastDeviceName = FirstNonBlank(target.LastDeviceName, source.LastDeviceName);
            target.LastRegistryKey = FirstNonBlank(target.LastRegistryKey, source.LastRegistryKey);
            target.LastSerialNumber = FirstNonBlank(target.LastSerialNumber, source.LastSerialNumber);
            target.LastInstanceId = FirstNonBlank(target.LastInstanceId, source.LastInstanceId);
            target.LastMonitorId = FirstNonBlank(target.LastMonitorId, source.LastMonitorId);

            foreach (var raw in source.KnownTargets)
            {
                var targetName = NormaliseTarget(raw);
                if (string.IsNullOrWhiteSpace(targetName)) continue;
                if (target.KnownTargets.Any(x => x.Equals(targetName, StringComparison.OrdinalIgnoreCase))) continue;
                target.KnownTargets.Add(targetName);
            }

            const int MaxKnownTargets = 8;
            if (target.KnownTargets.Count > MaxKnownTargets)
                target.KnownTargets.RemoveRange(MaxKnownTargets, target.KnownTargets.Count - MaxKnownTargets);
        }

        private static bool IsPlaceholderAlias(string? alias, string stableKey)
        {
            if (string.IsNullOrWhiteSpace(alias)) return true;
            var value = alias.Trim();
            return value.Equals(stableKey, StringComparison.OrdinalIgnoreCase) ||
                   value.StartsWith("DEV:", StringComparison.OrdinalIgnoreCase);
        }

        private static string? FirstNonBlank(string? current, string? fallback)
            => !string.IsNullOrWhiteSpace(current) ? current : fallback;

        private void EnsureOnlyPreferredPrimary(string stableKey)
        {
            foreach (var key in _aliasMap.Keys.ToList())
            {
                var info = _aliasMap[key];
                info.IsPreferredPrimary = key.Equals(stableKey, StringComparison.OrdinalIgnoreCase);
                _aliasMap[key] = info;
            }
        }

        private void EnsureOnlyFallbackPrimary(string stableKey)
        {
            foreach (var key in _aliasMap.Keys.ToList())
            {
                var info = _aliasMap[key];
                info.IsFallbackPrimary = key.Equals(stableKey, StringComparison.OrdinalIgnoreCase);
                _aliasMap[key] = info;
            }
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

        private void LogDetectionSnapshotIfChanged()
        {
            var rows = _detected
                .OrderBy(d => d.PositionX)
                .ThenBy(d => d.DeviceName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(d => d.StableKey, StringComparer.OrdinalIgnoreCase)
                .Select(d => string.Join(", ",
                    $"key={LogValue(d.StableKey)}",
                    $"device={LogValue(d.DeviceName)}",
                    $"active={d.IsActive}",
                    $"present={d.IsPresent}",
                    $"x={d.PositionX}",
                    $"serial={LogValue(d.SerialNumber)}",
                    $"monitorId={LogValue(d.MonitorId)}",
                    $"instanceId={LogValue(d.InstanceId)}",
                    $"monitorKey={LogValue(d.MonitorKey)}"))
                .ToList();

            var signature = string.Join("|", rows);
            if (signature.Equals(_lastDetectionLogSignature, StringComparison.Ordinal))
                return;

            _lastDetectionLogSignature = signature;

            int active = _detected.Count(d => d.IsPresent && d.IsActive);
            int present = _detected.Count(d => d.IsPresent);
            _log.Write($"Detected monitor state changed: active={active}, present={present}, rows={_detected.Count}.");

            foreach (var row in rows)
                _log.Write($"Detected monitor: {row}.");
        }

        private static string LogValue(string? value)
            => string.IsNullOrWhiteSpace(value) ? "<blank>" : value.Trim();

        private static void LogMonitorAction(string message)
        {
            Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
        }

        private void CheckToolExists()
        {
            if (!File.Exists(_toolPath))
            {
                _log.Write($"MultiMonitorTool.exe not found at {_toolPath}. Detection will use Screen.AllScreens fallback; monitor actions require the tool.");
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
            var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
            if (bounds.Width > 0 && bounds.Height > 0)
            {
                _uiSettings.WindowX = bounds.X;
                _uiSettings.WindowY = bounds.Y;
                _uiSettings.WindowWidth = bounds.Width;
                _uiSettings.WindowHeight = bounds.Height;
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
