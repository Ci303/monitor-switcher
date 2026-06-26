#nullable enable
using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using WorkMonitorSwitcher.Model;
using WorkMonitorSwitcher.Services;

namespace WorkMonitorSwitcher
{
    /// <summary>
    /// Settings dialog: shows StableKey (short), RegistryKey and editable Alias.
    /// Also includes a Dark Mode checkbox that feeds back to Form1.
    /// </summary>
    public sealed class AliasSettingsForm : Form
    {
        private readonly DataGridView _grid = new()
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToResizeRows = false,
            ReadOnly = false,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        };

        private readonly Button _ok = new() { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
        private readonly Button _cancel = new() { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
        private readonly Button _remove = new() { Text = "Remove Selected", AutoSize = true };
        private readonly Button _openRegistry = new() { Text = "Open Registry", AutoSize = true };
        private readonly Button _editCfg = new() { Text = "Edit cfg", AutoSize = true };
        private readonly Button _downloadTool = new() { Text = "Download MultiMonitorTool", AutoSize = true };
        private readonly Button _updateApp = new() { Text = "Update App", AutoSize = true };
        private readonly Button _showDiagnostics = new() { Text = "Diagnostics", AutoSize = true };
        private readonly CheckBox _chkDark = new() { Text = "Dark mode", AutoSize = true };
        private readonly CheckBox _chkTopMost = new() { Text = "Always on top", AutoSize = true };
        private readonly CheckBox _chkTray = new() { Text = "Minimize to tray", AutoSize = true };
        private readonly CheckBox _chkStartup = new() { Text = "Start with Windows", AutoSize = true };
        private readonly CheckBox _chkConfirmDisable = new() { Text = "Confirm before disabling", AutoSize = true };
        private readonly Button _layoutProfileButton = new() { Text = "Default", AutoSize = false, Size = new Size(150, 26), TextAlign = ContentAlignment.MiddleLeft };
        private readonly Button _deleteProfile = new() { Text = "Delete Profile", AutoSize = false, Size = new Size(96, 26) };
        private readonly TextBox _details = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
        private readonly BindingList<AliasViewRow> _rows;
        private readonly string _cfgPath;
        private readonly string _diagnosticsText;
        private readonly List<string> _layoutProfileNames;
        private string _selectedLayoutProfile;
        private ContextMenuStrip? _layoutProfileMenu;
        private readonly ToolTip _toolTip = new() { InitialDelay = 350, ReshowDelay = 100, AutoPopDelay = 8000 };
        private static readonly HttpClient Http = new();
        private const string MultiMonitorToolZipUrl = "https://www.nirsoft.net/utils/multimonitortool-x64.zip";
        private const string LatestReleaseApiUrl = "https://api.github.com/repos/Ci303/monitor-switcher/releases/latest";

        public Dictionary<string, string> UpdatedMappings { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> RemovedKeys { get; } = new(StringComparer.OrdinalIgnoreCase);
        public bool? DarkModeResult { get; private set; }
        public string? PreferredPrimaryKey { get; private set; }
        public string? FallbackPrimaryKey { get; private set; }
        public bool? AlwaysOnTopResult { get; private set; }
        public bool? MinimizeToTrayResult { get; private set; }
        public bool? StartWithWindowsResult { get; private set; }
        public bool? ConfirmBeforeDisableResult { get; private set; }
        public string? SelectedLayoutProfileResult { get; private set; }
        public HashSet<string> RemovedLayoutProfiles { get; } = new(StringComparer.OrdinalIgnoreCase);

        public AliasSettingsForm(
            List<AliasViewRow> current,
            bool darkMode,
            bool alwaysOnTop,
            bool minimizeToTray,
            bool startWithWindows,
            bool confirmBeforeDisable,
            string cfgPath,
            List<string> layoutProfiles,
            string selectedLayoutProfile,
            string diagnosticsText,
            Form? sizingOwner = null)

        {
            _rows = new BindingList<AliasViewRow>(current);
            _cfgPath = cfgPath ?? string.Empty;
            _diagnosticsText = string.IsNullOrWhiteSpace(diagnosticsText)
                ? "No diagnostic events have been recorded yet."
                : diagnosticsText;
            _layoutProfileNames = layoutProfiles
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(p => p.Equals("Default", StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                .ThenBy(p => p, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (_layoutProfileNames.Count == 0)
                _layoutProfileNames.Add("Default");
            _selectedLayoutProfile = string.IsNullOrWhiteSpace(selectedLayoutProfile) ? "Default" : selectedLayoutProfile.Trim();
            if (!_layoutProfileNames.Any(p => p.Equals(_selectedLayoutProfile, StringComparison.OrdinalIgnoreCase)))
                _layoutProfileNames.Add(_selectedLayoutProfile);

            Text = "Monitor Switcher Settings";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            ShowIcon = true;
            try
            {
                Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch
            {
                // Non-fatal: the dialog can still open without a title bar icon.
            }
            MinimumSize = new Size(920, 430);
            Size = new Size(1120, 660);
            _chkDark.Checked = darkMode;
            _chkTopMost.Checked = alwaysOnTop;
            _chkTray.Checked = minimizeToTray;
            _chkStartup.Checked = startWithWindows;
            _chkConfirmDisable.Checked = confirmBeforeDisable;
            SetSelectedLayoutProfile(_selectedLayoutProfile);
            _grid.RowHeadersVisible = false;
            _grid.AllowUserToOrderColumns = true;
            _grid.MultiSelect = true;
            _grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            _grid.BorderStyle = BorderStyle.None;
            _grid.RowTemplate.Height = 28;
            _grid.MinimumSize = new Size(700, 0);
            _details.MinimumSize = new Size(320, 0);

            // Columns
            var shortKeyCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "Stable Key",
                DataPropertyName = nameof(AliasViewRow.ShortKey),
                ReadOnly = true,
                FillWeight = 28,
                MinimumWidth = 150
            };
            var regCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "Registry Key",
                DataPropertyName = nameof(AliasViewRow.RegistryKeyShort),
                ReadOnly = true,
                FillWeight = 44,
                MinimumWidth = 220
            };
            var aliasCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "Alias",
                DataPropertyName = nameof(AliasViewRow.Alias),
                ReadOnly = false,
                FillWeight = 28,
                MinimumWidth = 160
            };
            var primaryCol = new DataGridViewCheckBoxColumn

            {
                HeaderText = "Primary",
                DataPropertyName = nameof(AliasViewRow.IsPreferredPrimary),
                ReadOnly = false,
                FillWeight = 10,
                MinimumWidth = 70
            };

            primaryCol.ThreeState = false;
            var fallbackPrimaryCol = new DataGridViewCheckBoxColumn
            {
                HeaderText = "Fallback",
                DataPropertyName = nameof(AliasViewRow.IsFallbackPrimary),
                ReadOnly = false,
                FillWeight = 10,
                MinimumWidth = 76
            };
            fallbackPrimaryCol.ThreeState = false;

            _grid.Columns.Add(shortKeyCol);
            _grid.Columns.Add(regCol);
            _grid.Columns.Add(aliasCol);
            _grid.Columns.Add(primaryCol);
            _grid.Columns.Add(fallbackPrimaryCol);
            _grid.DataSource = _rows;
            _grid.EditMode = DataGridViewEditMode.EditOnEnter;
            _grid.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (_grid.IsCurrentCellDirty)
                    _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            // Tooltip with full StableKey when hovering the first column
            _grid.CellFormatting += (_, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == 0)
                {
                    var row = _rows[e.RowIndex];
                    _grid.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = row.StableKey;
                }
                else if (e.RowIndex >= 0 &&
                         e.ColumnIndex >= 0 &&
                         _grid.Columns[e.ColumnIndex].DataPropertyName == nameof(AliasViewRow.RegistryKeyShort))
                {
                    var row = _rows[e.RowIndex];
                    _grid.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = row.RegistryKey;
                }
            };

            // Ctrl+C copies current cell text
            _grid.KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.C && _grid.CurrentCell != null)
                {
                    var text = _grid.Columns[_grid.CurrentCell.ColumnIndex].DataPropertyName == nameof(AliasViewRow.RegistryKeyShort)
                        ? _rows[_grid.CurrentCell.RowIndex].RegistryKey
                        : Convert.ToString(_grid.CurrentCell.Value) ?? string.Empty;
                    if (!string.IsNullOrEmpty(text))
                        Clipboard.SetText(text);
                    e.Handled = true;
                }
            };

            _grid.CellValueChanged += (s, e) =>
            {
                if (e.RowIndex < 0) return;

                if (_grid.Columns[e.ColumnIndex] is not DataGridViewCheckBoxColumn)
                    return;

                var propertyName = _grid.Columns[e.ColumnIndex].DataPropertyName;
                if (propertyName == nameof(AliasViewRow.IsPreferredPrimary))
                {
                    KeepSingleCheckedRow(
                        e.RowIndex,
                        row => row.IsPreferredPrimary,
                        (row, value) => row.IsPreferredPrimary = value);
                }
                else if (propertyName == nameof(AliasViewRow.IsFallbackPrimary))
                {
                    KeepSingleCheckedRow(
                        e.RowIndex,
                        row => row.IsFallbackPrimary,
                        (row, value) => row.IsFallbackPrimary = value);
                }
            };
            _grid.CellDoubleClick += (_, e) =>
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                if (_grid.Columns[e.ColumnIndex].DataPropertyName == nameof(AliasViewRow.RegistryKeyShort))
                    OpenRegistryForRow(e.RowIndex);
            };
            _grid.SelectionChanged += (_, __) => UpdateDetailsPanel();
            _remove.Click += (_, __) => RemoveSelectedRows();
            _openRegistry.Click += (_, __) => OpenRegistryForSelectedRow();
            _layoutProfileButton.Click += (_, __) => ShowLayoutProfileMenu();
            _deleteProfile.Click += (_, __) => DeleteSelectedProfile();
            _editCfg.Click += (_, __) => ChooseAndOpenCfg();
            _downloadTool.Click += async (_, __) => await DownloadToolAsync();
            _updateApp.Click += async (_, __) => await UpdateAppAsync();
            _showDiagnostics.Click += (_, __) => ShowDiagnosticsDialog();
            _toolTip.SetToolTip(_chkDark, "Use the dark color theme.");
            _toolTip.SetToolTip(_chkTopMost, "Keep the main switcher window above other windows.");
            _toolTip.SetToolTip(_chkTray, "Close to the notification area instead of exiting.");
            _toolTip.SetToolTip(_chkStartup, "Start Monitor Switcher when you sign in to Windows.");
            _toolTip.SetToolTip(_chkConfirmDisable, "Ask before disabling a monitor.");
            _toolTip.SetToolTip(_layoutProfileButton, "Current layout profile used by the main window.");
            _toolTip.SetToolTip(_deleteProfile, "Delete the selected saved layout profile. Default cannot be deleted.");
            _toolTip.SetToolTip(_remove, "Remove selected saved monitor entries.");
            _toolTip.SetToolTip(_openRegistry, "Open Registry Editor at the selected monitor key.");
            _toolTip.SetToolTip(_editCfg, "Open MultiMonitorTool.cfg.");
            _toolTip.SetToolTip(_downloadTool, "Download the official NirSoft helper beside the app.");
            _toolTip.SetToolTip(_updateApp, "Download the latest GitHub release asset if one is published.");
            _toolTip.SetToolTip(_showDiagnostics, "Show recent monitor action and layout profile events.");

            // Top bar
            var topBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 44,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(12, 10, 12, 6),
                AutoSize = false,
                WrapContents = false,
                AutoScroll = true
            };
            _chkDark.Margin = new Padding(0, 3, 18, 3);
            _chkTopMost.Margin = new Padding(0, 3, 18, 3);
            _chkTray.Margin = new Padding(0, 3, 18, 3);
            _chkStartup.Margin = new Padding(0, 3, 18, 3);
            _chkConfirmDisable.Margin = new Padding(0, 3, 18, 3);
            topBar.Controls.Add(_chkDark);
            topBar.Controls.Add(_chkTopMost);
            topBar.Controls.Add(_chkTray);
            topBar.Controls.Add(_chkStartup);
            topBar.Controls.Add(_chkConfirmDisable);

            var profileBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 42,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(12, 7, 12, 5),
                AutoScroll = true
            };
            var profileLabel = new Label
            {
                Text = "Layout profile",
                AutoSize = true,
                Margin = new Padding(0, 5, 8, 0)
            };
            _layoutProfileButton.Margin = new Padding(0, 0, 8, 0);
            _deleteProfile.Margin = new Padding(0, 0, 8, 0);
            profileBar.Controls.Add(profileLabel);
            profileBar.Controls.Add(_layoutProfileButton);
            profileBar.Controls.Add(_deleteProfile);

            var bottom = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(12, 8, 12, 10),
                Height = 54
            };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var toolActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = true,
                AutoScroll = true
            };
            var commitActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = true
            };

            foreach (var button in new[] { _editCfg, _openRegistry, _downloadTool, _updateApp, _showDiagnostics, _remove, _cancel, _ok })
                button.Margin = new Padding(0, 0, 8, 0);

            toolActions.Controls.Add(_editCfg);
            toolActions.Controls.Add(_openRegistry);
            toolActions.Controls.Add(_downloadTool);
            toolActions.Controls.Add(_updateApp);
            toolActions.Controls.Add(_showDiagnostics);
            commitActions.Controls.Add(_ok);
            commitActions.Controls.Add(_cancel);
            commitActions.Controls.Add(_remove);
            bottom.Controls.Add(toolActions, 0, 0);
            bottom.Controls.Add(commitActions, 1, 0);

            var monitorLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(8, 8, 8, 0)
            };
            monitorLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 72));
            monitorLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28));
            monitorLayout.Controls.Add(_grid, 0, 0);
            monitorLayout.Controls.Add(_details, 1, 0);

            Controls.Add(monitorLayout);
            Controls.Add(topBar);
            Controls.Add(profileBar);
            Controls.Add(bottom);
            FitInitialSizeToContent(topBar, profileBar, toolActions, commitActions, sizingOwner);

            AcceptButton = _ok;
            CancelButton = _cancel;

            // Initial theme
            ApplyDialogTheme(darkMode);

            _chkDark.CheckedChanged += (s, e) => ApplyDialogTheme(_chkDark.Checked);

            _ok.Click += (_, __) =>
            {
                _grid.EndEdit();
                UpdatedMappings.Clear();
                foreach (var row in _rows)
                    UpdatedMappings[row.StableKey] = row.Alias ?? string.Empty;

                DarkModeResult = _chkDark.Checked;
                PreferredPrimaryKey = _rows.FirstOrDefault(r => r.IsPreferredPrimary)?.StableKey;
                FallbackPrimaryKey = _rows.FirstOrDefault(r => r.IsFallbackPrimary)?.StableKey;
                AlwaysOnTopResult = _chkTopMost.Checked;
                MinimizeToTrayResult = _chkTray.Checked;
                StartWithWindowsResult = _chkStartup.Checked;
                ConfirmBeforeDisableResult = _chkConfirmDisable.Checked;
                SelectedLayoutProfileResult = _selectedLayoutProfile;
                DialogResult = DialogResult.OK;
                Close();
            };

            UpdateDetailsPanel();
        }

        private void FitInitialSizeToContent(
            FlowLayoutPanel topBar,
            FlowLayoutPanel profileBar,
            FlowLayoutPanel toolActions,
            FlowLayoutPanel commitActions,
            Form? sizingOwner)
        {
            int gridMinWidth = _grid.Columns
                .Cast<DataGridViewColumn>()
                .Sum(c => c.MinimumWidth) +
                SystemInformation.VerticalScrollBarWidth +
                16;

            int monitorAreaWidth = gridMinWidth + _details.MinimumSize.Width + 44;
            int topBarWidth = topBar.Padding.Horizontal + PreferredControlsWidth(topBar.Controls) + 24;
            int profileBarWidth = profileBar.Padding.Horizontal + PreferredControlsWidth(profileBar.Controls) + 24;
            int bottomWidth =
                PreferredControlsWidth(toolActions.Controls) +
                PreferredControlsWidth(commitActions.Controls) +
                72;

            int desiredClientWidth = Math.Max(
                1120,
                Math.Max(monitorAreaWidth, Math.Max(topBarWidth, Math.Max(profileBarWidth, bottomWidth))));

            int visibleRows = Math.Min(Math.Max(_rows.Count, 5), 9);
            int desiredGridHeight = _grid.ColumnHeadersHeight + (visibleRows * _grid.RowTemplate.Height) + 76;
            int desiredClientHeight = Math.Max(
                640,
                topBar.Height + profileBar.Height + 54 + desiredGridHeight + 32);

            var screen = sizingOwner != null && !sizingOwner.IsDisposed
                ? Screen.FromControl(sizingOwner)
                : Screen.FromPoint(Cursor.Position);
            var workingArea = screen.WorkingArea;
            int maxClientWidth = Math.Max(640, workingArea.Width - 80);
            int maxClientHeight = Math.Max(480, workingArea.Height - 80);

            var desiredClientSize = new Size(
                Math.Min(desiredClientWidth, maxClientWidth),
                Math.Min(desiredClientHeight, maxClientHeight));

            var minimumClientSize = new Size(
                Math.Min(920, maxClientWidth),
                Math.Min(560, maxClientHeight));

            MinimumSize = SizeFromClientSize(minimumClientSize);
            ClientSize = desiredClientSize;
        }

        private void KeepSingleCheckedRow(
            int selectedRowIndex,
            Func<AliasViewRow, bool> isChecked,
            Action<AliasViewRow, bool> setChecked)
        {
            if (!isChecked(_rows[selectedRowIndex]))
                return;

            for (int i = 0; i < _rows.Count; i++)
            {
                if (i != selectedRowIndex && isChecked(_rows[i]))
                    setChecked(_rows[i], false);
            }

            _grid.Invalidate();
        }

        private static int PreferredControlsWidth(Control.ControlCollection controls)
        {
            int width = 0;
            foreach (Control control in controls)
            {
                var preferred = control.GetPreferredSize(Size.Empty);
                width += Math.Max(control.Width, preferred.Width) + control.Margin.Horizontal;
            }

            return width;
        }

        private void RemoveSelectedRows()
        {
            var selectedRows = _grid.SelectedCells
                .Cast<DataGridViewCell>()
                .Select(c => c.RowIndex)
                .Where(i => i >= 0)
                .Distinct()
                .OrderByDescending(i => i)
                .ToList();

            if (selectedRows.Count == 0)
                return;

            if (MessageBox.Show(
                    this,
                    "Remove the selected monitor entries from saved settings? Connected monitors will reappear after the next refresh.",
                    "Remove Monitor Entries",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            foreach (var rowIndex in selectedRows)
            {
                RemovedKeys.Add(_rows[rowIndex].StableKey);
                _rows.RemoveAt(rowIndex);
            }
            UpdateDetailsPanel();
        }

        private void DeleteSelectedProfile()
        {
            var profile = (_selectedLayoutProfile ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(profile) || profile.Equals("Default", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this, "The Default layout profile cannot be deleted.", "Delete Profile",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show(this, $"Delete layout profile '{profile}'?", "Delete Profile",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            RemovedLayoutProfiles.Add(profile);
            _layoutProfileNames.RemoveAll(p => p.Equals(profile, StringComparison.OrdinalIgnoreCase));
            SetSelectedLayoutProfile(_layoutProfileNames.FirstOrDefault() ?? "Default");
        }

        private void SetSelectedLayoutProfile(string profile)
        {
            _selectedLayoutProfile = string.IsNullOrWhiteSpace(profile) ? "Default" : profile.Trim();
            _layoutProfileButton.Text = _selectedLayoutProfile;
            BuildLayoutProfileMenu();
        }

        private void BuildLayoutProfileMenu()
        {
            _layoutProfileMenu?.Dispose();

            var palette = _chkDark.Checked ? ThemePalette.Dark() : ThemePalette.Light();
            _layoutProfileMenu = new ContextMenuStrip
            {
                BackColor = palette.Back,
                ForeColor = palette.Text,
                Renderer = new LayoutProfileMenuRenderer(palette),
                ShowCheckMargin = true,
                ShowImageMargin = false
            };

            _layoutProfileMenu.BackColor = palette.Surface;
            _layoutProfileMenu.ForeColor = palette.Text;

            foreach (var profile in _layoutProfileNames
                         .OrderBy(p => p.Equals("Default", StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                         .ThenBy(p => p, StringComparer.OrdinalIgnoreCase))
            {
                var item = new ToolStripMenuItem(profile)
                {
                    Checked = profile.Equals(_selectedLayoutProfile, StringComparison.OrdinalIgnoreCase),
                    ForeColor = palette.Text
                };
                item.Click += (_, __) => SetSelectedLayoutProfile(profile);
                _layoutProfileMenu.Items.Add(item);
            }
        }

        private void ShowLayoutProfileMenu()
        {
            BuildLayoutProfileMenu();
            _layoutProfileMenu?.Show(_layoutProfileButton, new Point(0, _layoutProfileButton.Height));
        }

        private void ShowDiagnosticsDialog()
        {
            try
            {
                var path = Path.Combine(Path.GetTempPath(), "MonitorSwitcher-diagnostics.txt");
                File.WriteAllText(path, _diagnosticsText);
                Process.Start(new ProcessStartInfo("notepad.exe", $"\"{path}\"") { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Unable to open diagnostics.\n{ex.Message}", "Diagnostics",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void UpdateDetailsPanel()
        {
            var rowIndex = _grid.CurrentCell?.RowIndex ?? -1;
            if (rowIndex < 0 || rowIndex >= _rows.Count)
            {
                _details.Text = "Select a monitor to view saved identity details.";
                return;
            }

            var row = _rows[rowIndex];
            _details.Text =
                $"Alias: {row.Alias}{Environment.NewLine}" +
                $"Preferred primary: {(row.IsPreferredPrimary ? "Yes" : "No")}{Environment.NewLine}" +
                $"Fallback primary: {(row.IsFallbackPrimary ? "Yes" : "No")}{Environment.NewLine}" +
                $"{Environment.NewLine}" +
                $"Stable key:{Environment.NewLine}{row.StableKeyFull}{Environment.NewLine}" +
                $"{Environment.NewLine}" +
                $"Registry key:{Environment.NewLine}{row.RegistryKey}{Environment.NewLine}" +
                $"{Environment.NewLine}" +
                $"Device name: {row.DeviceName}{Environment.NewLine}" +
                $"Monitor name: {row.MonitorName}{Environment.NewLine}" +
                $"Monitor ID: {row.MonitorId}{Environment.NewLine}" +
                $"Instance ID: {row.InstanceId}{Environment.NewLine}" +
                $"Serial number: {row.SerialNumber}{Environment.NewLine}" +
                $"{Environment.NewLine}" +
                $"Known command targets:{Environment.NewLine}{row.KnownTargets}";
        }

        private void OpenRegistryForSelectedRow()
        {
            var rowIndex = _grid.CurrentCell?.RowIndex ?? -1;
            if (rowIndex < 0 && _grid.SelectedCells.Count > 0)
                rowIndex = _grid.SelectedCells.Cast<DataGridViewCell>().Min(c => c.RowIndex);

            OpenRegistryForRow(rowIndex);
        }

        private void OpenRegistryForRow(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= _rows.Count)
                return;

            var rawPath = _rows[rowIndex].RegistryKey;
            if (string.IsNullOrWhiteSpace(rawPath))
            {
                MessageBox.Show(this, "This monitor does not have a saved registry key yet.", "Open Registry",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var regeditPath = ToRegeditLastKeyPath(rawPath);
            if (string.IsNullOrWhiteSpace(regeditPath))
            {
                MessageBox.Show(this, $"Unable to open this registry path:\n{rawPath}", "Open Registry",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Applets\Regedit"))
                    key?.SetValue("LastKey", regeditPath, RegistryValueKind.String);

                Process.Start(new ProcessStartInfo("regedit.exe") { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Unable to open Registry Editor.\n{ex.Message}", "Open Registry",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static string ToRegeditLastKeyPath(string rawPath)
        {
            var path = rawPath.Trim();
            const string machinePrefix = @"\Registry\Machine\";
            const string userPrefix = @"\Registry\User\";

            if (path.StartsWith(machinePrefix, StringComparison.OrdinalIgnoreCase))
                return @"Computer\HKEY_LOCAL_MACHINE\" + path[machinePrefix.Length..];

            if (path.StartsWith(userPrefix, StringComparison.OrdinalIgnoreCase))
                return @"Computer\HKEY_USERS\" + path[userPrefix.Length..];

            if (path.StartsWith(@"HKEY_", StringComparison.OrdinalIgnoreCase))
                return @"Computer\" + path;

            if (path.StartsWith(@"Computer\HKEY_", StringComparison.OrdinalIgnoreCase))
                return path;

            return string.Empty;
        }

        private void ChooseAndOpenCfg()
        {
            var choice = MessageBox.Show(
                this,
                "Open MultiMonitorTool.cfg with the default editor?\n\nChoose Yes for the default editor or No for Notepad.",
                "Edit cfg",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (choice == DialogResult.Cancel)
                return;

            OpenCfg(preferNotepad: choice == DialogResult.No);
        }

        private void OpenCfg(bool preferNotepad)
        {
            try
            {
                EnsureCfgFileExists();

                var psi = preferNotepad
                    ? new ProcessStartInfo("notepad.exe", $"\"{_cfgPath}\"")
                    : new ProcessStartInfo(_cfgPath) { UseShellExecute = true };

                if (!preferNotepad)
                    psi.UseShellExecute = true;

                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Unable to open cfg file.\n{ex.Message}", "Monitor Aliases",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private async Task DownloadToolAsync()
        {
            string exePath = Path.Combine(Path.GetDirectoryName(_cfgPath) ?? AppContext.BaseDirectory, "MultiMonitorTool.exe");

            var overwrite = true;
            if (File.Exists(exePath))
            {
                var choice = MessageBox.Show(
                    this,
                    "MultiMonitorTool.exe already exists beside the app. Replace it with the latest official x64 version from NirSoft?",
                    "Download MultiMonitorTool",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (choice != DialogResult.Yes)
                    overwrite = false;
            }

            if (!overwrite)
                return;

            var originalText = _downloadTool.Text;
            Enabled = false;
            UseWaitCursor = true;
            _downloadTool.Text = "Downloading...";

            string zipPath = Path.Combine(Path.GetTempPath(), $"multimonitortool-{Guid.NewGuid():N}.zip");
            string extractDir = Path.Combine(Path.GetTempPath(), $"multimonitortool-{Guid.NewGuid():N}");

            try
            {
                Directory.CreateDirectory(extractDir);

                var bytes = await Http.GetByteArrayAsync(MultiMonitorToolZipUrl);
                await File.WriteAllBytesAsync(zipPath, bytes);
                ZipFile.ExtractToDirectory(zipPath, extractDir, overwriteFiles: true);

                string sourceExe = Path.Combine(extractDir, "MultiMonitorTool.exe");
                if (!File.Exists(sourceExe))
                    throw new FileNotFoundException("Downloaded package did not contain MultiMonitorTool.exe.", sourceExe);

                Directory.CreateDirectory(Path.GetDirectoryName(exePath)!);
                File.Copy(sourceExe, exePath, overwrite: true);

                MessageBox.Show(
                    this,
                    $"MultiMonitorTool.exe downloaded successfully to:\n{exePath}\n\n" +
                    "Keep MultiMonitorTool.cfg in the same folder as MultiMonitorTool.exe and MonitorSwitcher.exe.",
                    "Download MultiMonitorTool",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    $"Unable to download MultiMonitorTool.\n{ex.Message}",
                    "Download MultiMonitorTool",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                try
                {
                    if (File.Exists(zipPath))
                        File.Delete(zipPath);
                }
                catch { /* ignore */ }

                try
                {
                    if (Directory.Exists(extractDir))
                        Directory.Delete(extractDir, recursive: true);
                }
                catch { /* ignore */ }

                _downloadTool.Text = originalText;
                UseWaitCursor = false;
                Enabled = true;
            }
        }

        private async Task UpdateAppAsync()
        {
            var originalText = _updateApp.Text;
            Enabled = false;
            UseWaitCursor = true;
            _updateApp.Text = "Checking...";

            try
            {
                using var request = new HttpRequestMessage(HttpMethod.Get, LatestReleaseApiUrl);
                request.Headers.UserAgent.Add(new ProductInfoHeaderValue("MonitorSwitcher", "1.0"));
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));

                using var response = await Http.SendAsync(request);
                if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    MessageBox.Show(
                        this,
                        "No GitHub release is published for this repository yet.",
                        "Update App",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                response.EnsureSuccessStatusCode();

                await using var stream = await response.Content.ReadAsStreamAsync();
                var release = await JsonSerializer.DeserializeAsync<GitHubRelease>(stream);
                if (release == null)
                    throw new InvalidOperationException("Unable to read the latest release details.");

                var asset = release.Assets
                    .FirstOrDefault(a =>
                        a.Name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) ||
                        a.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase));

                if (asset == null || string.IsNullOrWhiteSpace(asset.DownloadUrl))
                {
                    var choice = MessageBox.Show(
                        this,
                        "The latest release has no downloadable .zip or .exe asset.\n\nOpen the release page instead?",
                        "Update App",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (choice == DialogResult.Yes && !string.IsNullOrWhiteSpace(release.HtmlUrl))
                        Process.Start(new ProcessStartInfo(release.HtmlUrl) { UseShellExecute = true });

                    return;
                }

                _updateApp.Text = "Downloading...";

                string appDir = Path.GetDirectoryName(_cfgPath) ?? AppContext.BaseDirectory;
                string updatesDir = Path.Combine(appDir, "updates", SanitizePathSegment(release.TagName));
                Directory.CreateDirectory(updatesDir);

                string assetPath = Path.Combine(updatesDir, asset.Name);
                var bytes = await Http.GetByteArrayAsync(asset.DownloadUrl);
                await File.WriteAllBytesAsync(assetPath, bytes);

                string finalPath = assetPath;
                if (asset.Name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    string extractDir = Path.Combine(updatesDir, Path.GetFileNameWithoutExtension(asset.Name));
                    Directory.CreateDirectory(extractDir);
                    ZipFile.ExtractToDirectory(assetPath, extractDir, overwriteFiles: true);
                    finalPath = extractDir;
                }

                var openChoice = MessageBox.Show(
                    this,
                    $"Downloaded {asset.Name} from release {release.TagName}.\n\nSaved to:\n{finalPath}\n\nOpen the folder now?",
                    "Update App",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (openChoice == DialogResult.Yes)
                    Process.Start(new ProcessStartInfo("explorer.exe", $"\"{finalPath}\"") { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    $"Unable to update from GitHub releases.\n{ex.Message}",
                    "Update App",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                _updateApp.Text = originalText;
                UseWaitCursor = false;
                Enabled = true;
            }
        }

        private static string SanitizePathSegment(string? raw)
        {
            var value = string.IsNullOrWhiteSpace(raw) ? "latest" : raw.Trim();
            foreach (var c in Path.GetInvalidFileNameChars())
                value = value.Replace(c, '_');
            return value;
        }

        private sealed class GitHubRelease
        {
            [JsonPropertyName("tag_name")]
            public string TagName { get; set; } = string.Empty;

            [JsonPropertyName("html_url")]
            public string HtmlUrl { get; set; } = string.Empty;

            [JsonPropertyName("assets")]
            public List<GitHubReleaseAsset> Assets { get; set; } = new();
        }

        private sealed class GitHubReleaseAsset
        {
            [JsonPropertyName("name")]
            public string Name { get; set; } = string.Empty;

            [JsonPropertyName("browser_download_url")]
            public string DownloadUrl { get; set; } = string.Empty;
        }

        private void EnsureCfgFileExists()
        {
            if (string.IsNullOrWhiteSpace(_cfgPath))
                throw new InvalidOperationException("Cfg path is not configured.");

            var dir = Path.GetDirectoryName(_cfgPath);
            if (!string.IsNullOrWhiteSpace(dir))
                Directory.CreateDirectory(dir);

            if (!File.Exists(_cfgPath))
                File.WriteAllText(_cfgPath, "MonitorsConfigNumOfCalls=5" + Environment.NewLine);
        }

        private void ApplyDialogTheme(bool dark)
        {
            var palette = dark ? ThemePalette.Dark() : ThemePalette.Light();

            // Apply palette to this form and its controls
            Themer.Apply(this, palette);

            // Title bar
            DwmInterop.SetDarkTitleBar(this.Handle, dark);
            BuildLayoutProfileMenu();

            // We don’t set caption color here; leaving it to system accent avoids visual mismatch
            Invalidate(true);
            Update();
        }

        private sealed class LayoutProfileMenuRenderer : ToolStripProfessionalRenderer
        {
            private readonly ThemePalette _palette;

            public LayoutProfileMenuRenderer(ThemePalette palette)
                : base(new LayoutProfileColorTable(palette))
            {
                _palette = palette;
            }

            protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
            {
                var item = e.Item as ToolStripMenuItem;
                bool selected = e.Item.Selected;
                bool isChecked = item?.Checked == true;

                var bounds = new Rectangle(Point.Empty, e.Item.Size);
                var backColor = selected
                    ? ControlPaint.Light(_palette.Surface, 0.18f)
                    : isChecked
                        ? ControlPaint.Light(_palette.Surface, 0.08f)
                        : _palette.Back;

                using var back = new SolidBrush(backColor);
                e.Graphics.FillRectangle(back, bounds);

                if (selected || isChecked)
                {
                    using var border = new Pen(_palette.Border);
                    e.Graphics.DrawRectangle(border, 0, 0, bounds.Width - 1, bounds.Height - 1);
                }
            }

            protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
            {
                e.TextColor = _palette.Text;
                base.OnRenderItemText(e);
            }

            protected override void OnRenderImageMargin(ToolStripRenderEventArgs e)
            {
                using var back = new SolidBrush(_palette.Surface);
                e.Graphics.FillRectangle(back, e.AffectedBounds);
            }
        }

        private sealed class LayoutProfileColorTable : ProfessionalColorTable
        {
            private readonly ThemePalette _palette;
            private readonly Color _selected;
            private readonly Color _checked;

            public LayoutProfileColorTable(ThemePalette palette)
            {
                _palette = palette;
                _selected = ControlPaint.Light(palette.Surface, 0.18f);
                _checked = ControlPaint.Light(palette.Surface, 0.08f);
            }

            public override Color ToolStripDropDownBackground => _palette.Back;
            public override Color ImageMarginGradientBegin => _palette.Surface;
            public override Color ImageMarginGradientMiddle => _palette.Surface;
            public override Color ImageMarginGradientEnd => _palette.Surface;
            public override Color MenuBorder => _palette.Border;
            public override Color MenuItemBorder => _palette.Border;
            public override Color MenuItemSelected => _selected;
            public override Color MenuItemSelectedGradientBegin => _selected;
            public override Color MenuItemSelectedGradientEnd => _selected;
            public override Color CheckBackground => _checked;
            public override Color CheckPressedBackground => _selected;
            public override Color CheckSelectedBackground => _selected;
        }
    }
}
