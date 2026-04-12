#nullable enable
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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
        private readonly CheckBox _chkDark = new() { Text = "Dark mode", AutoSize = true };
        private readonly CheckBox _chkTopMost = new() { Text = "Always on top", AutoSize = true };

        public Dictionary<string, string> UpdatedMappings { get; } = new(StringComparer.OrdinalIgnoreCase);
        public bool? DarkModeResult { get; private set; }
        public string? PreferredPrimaryKey { get; private set; }
        public bool? AlwaysOnTopResult { get; private set; }



        public AliasSettingsForm(List<AliasViewRow> current, bool darkMode, bool alwaysOnTop)

        {
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
            _grid.DataSource = current;
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
                    var row = ((List<AliasViewRow>)_grid.DataSource)[e.RowIndex];
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
                    var rows = (List<AliasViewRow>)_grid.DataSource;

                    // Ensure single-select: if the clicked row is now true, set all others false
                    bool selected = rows[e.RowIndex].IsPreferredPrimary;
                    if (selected)
                    {
                        for (int i = 0; i < rows.Count; i++)
                            if (i != e.RowIndex && rows[i].IsPreferredPrimary)
                                rows[i].IsPreferredPrimary = false;

                        _grid.Invalidate(); // repaint checkboxes
                    }
                }
            };

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
            buttons.Controls.Add(_ok);
            buttons.Controls.Add(_cancel);

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
                foreach (var row in (List<AliasViewRow>)_grid.DataSource)
                    UpdatedMappings[row.StableKey] = row.Alias ?? string.Empty;

                DarkModeResult = _chkDark.Checked;
                var rowsOut = (List<AliasViewRow>)_grid.DataSource;
                PreferredPrimaryKey = rowsOut.FirstOrDefault(r => r.IsPreferredPrimary)?.StableKey;
                AlwaysOnTopResult = _chkTopMost.Checked;
                DialogResult = DialogResult.OK;
                Close();
            };
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
