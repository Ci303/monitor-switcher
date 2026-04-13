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
        private readonly Button _editCfg = new() { Text = "Edit cfg", AutoSize = true };
        private readonly Button _downloadTool = new() { Text = "Download MultiMonitorTool", AutoSize = true };
        private readonly Button _updateApp = new() { Text = "Update App", AutoSize = true };
        private readonly CheckBox _chkDark = new() { Text = "Dark mode", AutoSize = true };
        private readonly CheckBox _chkTopMost = new() { Text = "Always on top", AutoSize = true };
        private readonly BindingList<AliasViewRow> _rows;
        private readonly string _cfgPath;
        private static readonly HttpClient Http = new();
        private const string MultiMonitorToolZipUrl = "https://www.nirsoft.net/utils/multimonitortool-x64.zip";
        private const string LatestReleaseApiUrl = "https://api.github.com/repos/Ci303/monitor-switcher/releases/latest";

        public Dictionary<string, string> UpdatedMappings { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> RemovedKeys { get; } = new(StringComparer.OrdinalIgnoreCase);
        public bool? DarkModeResult { get; private set; }
        public string? PreferredPrimaryKey { get; private set; }
        public bool? AlwaysOnTopResult { get; private set; }



        public AliasSettingsForm(List<AliasViewRow> current, bool darkMode, bool alwaysOnTop, string cfgPath)

        {
            _rows = new BindingList<AliasViewRow>(current);
            _cfgPath = cfgPath ?? string.Empty;

            Text = "Monitor Aliases";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            MinimumSize = new Size(640, 360);
            Size = new Size(920, 560);
            _chkDark.Checked = darkMode;
            _chkTopMost.Checked = alwaysOnTop;

            // Columns
            var shortKeyCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "Stable Key (short)",
                DataPropertyName = nameof(AliasViewRow.ShortKey),
                ReadOnly = true,
                FillWeight = 28
            };
            var regCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "Registry Key",
                DataPropertyName = nameof(AliasViewRow.RegistryKey),
                ReadOnly = true,
                FillWeight = 47
            };
            var aliasCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "Alias",
                DataPropertyName = nameof(AliasViewRow.Alias),
                ReadOnly = false,
                FillWeight = 25
            };
            var primaryCol = new DataGridViewCheckBoxColumn

            {
                HeaderText = "Primary",
                DataPropertyName = nameof(AliasViewRow.IsPreferredPrimary),
                ReadOnly = false,
                FillWeight = 12
            };

            primaryCol.ThreeState = false;
            _grid.Columns.Add(shortKeyCol);
            _grid.Columns.Add(regCol);
            _grid.Columns.Add(aliasCol);
            _grid.Columns.Add(primaryCol);
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
            };

            // Ctrl+C copies current cell text
            _grid.KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.C && _grid.CurrentCell != null)
                {
                    var text = Convert.ToString(_grid.CurrentCell.Value) ?? string.Empty;
                    if (!string.IsNullOrEmpty(text))
                        Clipboard.SetText(text);
                    e.Handled = true;
                }
            };

            _grid.CellValueChanged += (s, e) =>
            {
                if (e.RowIndex < 0) return;

                // Only react to changes in the Primary column
                if (_grid.Columns[e.ColumnIndex] is DataGridViewCheckBoxColumn &&
                    _grid.Columns[e.ColumnIndex].DataPropertyName == nameof(AliasViewRow.IsPreferredPrimary))
                {
                    // Ensure single-select: if the clicked row is now true, set all others false
                    bool selected = _rows[e.RowIndex].IsPreferredPrimary;
                    if (selected)
                    {
                        for (int i = 0; i < _rows.Count; i++)
                            if (i != e.RowIndex && _rows[i].IsPreferredPrimary)
                                _rows[i].IsPreferredPrimary = false;

                        _grid.Invalidate(); // repaint checkboxes
                    }
                }
            };
            _remove.Click += (_, __) => RemoveSelectedRows();
            _editCfg.Click += (_, __) => ChooseAndOpenCfg();
            _downloadTool.Click += async (_, __) => await DownloadToolAsync();
            _updateApp.Click += async (_, __) => await UpdateAppAsync();

            // Top bar (dark mode toggle)
            var topBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 38,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(10, 8, 10, 4),
                AutoSize = false
            };
            topBar.Controls.Add(_chkDark);
            topBar.Controls.Add(_chkTopMost);

            // Bottom buttons
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                WrapContents = false,
                Padding = new Padding(10),
                Height = 52
            };
            _ok.Margin = new Padding(6, 6, 0, 6);
            _cancel.Margin = new Padding(6, 6, 6, 6);
            _remove.Margin = new Padding(6, 6, 6, 6);
            _editCfg.Margin = new Padding(6, 6, 6, 6);
            _downloadTool.Margin = new Padding(6, 6, 6, 6);
            _updateApp.Margin = new Padding(6, 6, 6, 6);
            buttons.Controls.Add(_ok);
            buttons.Controls.Add(_cancel);
            buttons.Controls.Add(_remove);
            buttons.Controls.Add(_editCfg);
            buttons.Controls.Add(_downloadTool);
            buttons.Controls.Add(_updateApp);

            Controls.Add(_grid);
            Controls.Add(topBar);
            Controls.Add(buttons);

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
                AlwaysOnTopResult = _chkTopMost.Checked;
                DialogResult = DialogResult.OK;
                Close();
            };
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

            // We don’t set caption color here; leaving it to system accent avoids visual mismatch
            Invalidate(true);
            Update();
        }
    }
}
