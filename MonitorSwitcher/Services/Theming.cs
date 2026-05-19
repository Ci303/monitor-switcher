using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WorkMonitorSwitcher.Services
{
    /// <summary>
    /// Simple palette for light/dark themes.
    /// </summary>
    internal sealed class ThemePalette
    {
        public Color Back { get; init; } = SystemColors.Window;
        public Color Surface { get; init; } = SystemColors.Control;
        public Color Border { get; init; } = Color.FromArgb(200, 200, 200);
        public Color Text { get; init; } = SystemColors.ControlText;
        public Color TextSubtle { get; init; } = Color.DimGray;
        public Color StatusOk { get; init; } = Color.Green;
        public Color StatusWarn { get; init; } = Color.Red;
        public Color StatusBusy { get; init; } = Color.DarkOrange;

        public static ThemePalette Light() => new();

        public static ThemePalette Dark() => new ThemePalette
        {
            Back = Color.FromArgb(30, 30, 30),
            Surface = Color.FromArgb(44, 44, 44),
            Border = Color.FromArgb(78, 78, 78),
            Text = Color.Gainsboro,
            TextSubtle = Color.Silver,
            StatusOk = Color.FromArgb(100, 200, 120),
            StatusWarn = Color.FromArgb(230, 100, 100),
            StatusBusy = Color.FromArgb(245, 170, 70)
        };
    }

    /// <summary>
    /// Applies a ThemePalette to a Form and its entire control tree.
    /// </summary>
    internal static class Themer
    {
        public static void Apply(Form form, ThemePalette p)
        {
            if (form == null) return;

            form.BackColor = p.Back;
            form.ForeColor = p.Text;

            foreach (Control c in form.Controls)
                ApplyToControl(c, p);

            form.Invalidate(true);
            form.Update();
        }

        private static void ApplyToControl(Control c, ThemePalette p)
        {
            switch (c)
            {
                case Label lbl:
                    lbl.BackColor = lbl.BorderStyle == BorderStyle.None ? Color.Transparent : p.Back;
                    // Preserve semantic colors if using our status tokens
                    if (lbl.Text.Equals("ONLINE", StringComparison.OrdinalIgnoreCase))
                        lbl.ForeColor = p.StatusOk;
                    else if (lbl.Text.Equals("DISABLED", StringComparison.OrdinalIgnoreCase))
                        lbl.ForeColor = p.StatusWarn;
                    else if (lbl.Text.Equals("OFFLINE", StringComparison.OrdinalIgnoreCase))
                        lbl.ForeColor = p.TextSubtle;
                    else if (lbl.Text.Equals("WORKING…", StringComparison.OrdinalIgnoreCase) || lbl.Text.Equals("WORKING...", StringComparison.OrdinalIgnoreCase))
                        lbl.ForeColor = p.StatusBusy;
                    else
                        lbl.ForeColor = p.Text;
                    break;

                case Button btn:
                    btn.BackColor = p.Surface;
                    btn.ForeColor = p.Text;
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.FlatAppearance.BorderColor = p.Border;
                    btn.FlatAppearance.BorderSize = 1;
                    btn.FlatAppearance.MouseOverBackColor = ControlPaint.Light(p.Surface, 0.06f);
                    btn.FlatAppearance.MouseDownBackColor = ControlPaint.Light(p.Surface, 0.12f);
                    btn.UseVisualStyleBackColor = false;
                    break;

                case DataGridView dgv:
                    dgv.BackColor = p.Back;
                    dgv.BackgroundColor = p.Back;
                    dgv.GridColor = p.Border;
                    dgv.EnableHeadersVisualStyles = false;

                    dgv.ColumnHeadersDefaultCellStyle.BackColor = p.Surface;
                    dgv.ColumnHeadersDefaultCellStyle.ForeColor = p.Text;
                    dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = p.Surface;
                    dgv.ColumnHeadersDefaultCellStyle.SelectionForeColor = p.Text;

                    dgv.DefaultCellStyle.BackColor = p.Back;
                    dgv.DefaultCellStyle.ForeColor = p.Text;
                    dgv.DefaultCellStyle.SelectionBackColor = ControlPaint.Light(p.Surface, 0.10f);
                    dgv.DefaultCellStyle.SelectionForeColor = p.Text;
                    dgv.AlternatingRowsDefaultCellStyle.BackColor = ControlPaint.Light(p.Back, 0.04f);
                    dgv.AlternatingRowsDefaultCellStyle.ForeColor = p.Text;

                    dgv.RowHeadersDefaultCellStyle.BackColor = p.Surface;
                    dgv.RowHeadersDefaultCellStyle.ForeColor = p.Text;
                    dgv.RowHeadersDefaultCellStyle.SelectionBackColor = p.Surface;
                    dgv.RowHeadersDefaultCellStyle.SelectionForeColor = p.Text;
                    EnsureNativeScrollTheme(dgv);
                    break;

                case ComboBox combo:
                    combo.BackColor = p.Back;
                    combo.ForeColor = p.Text;
                    combo.FlatStyle = FlatStyle.Flat;
                    combo.DrawMode = DrawMode.OwnerDrawFixed;
                    combo.DrawItem -= ComboDrawItem;
                    combo.DrawItem += ComboDrawItem;
                    combo.Tag = p;
                    break;

                case TextBox txt:
                    txt.BackColor = p.Back;
                    txt.ForeColor = p.Text;
                    txt.BorderStyle = BorderStyle.FixedSingle;
                    EnsureNativeScrollTheme(txt);
                    break;

                case TabControl tab:
                    tab.BackColor = p.Back;
                    tab.ForeColor = p.Text;
                    tab.DrawMode = TabDrawMode.OwnerDrawFixed;
                    tab.DrawItem -= TabDrawItem;
                    tab.DrawItem += TabDrawItem;
                    tab.Tag = p;
                    break;

                case TabPage page:
                    page.BackColor = p.Back;
                    page.ForeColor = p.Text;
                    break;

                case Panel pnl:
                    // Let 1px rule lines keep their custom color; other panels follow background
                    if (pnl.Height > 2)
                        pnl.BackColor = pnl.BorderStyle == BorderStyle.None ? p.Back : p.Surface;
                    break;

                default:
                    c.BackColor = p.Back;
                    c.ForeColor = p.Text;
                    break;
            }

            foreach (Control child in c.Controls)
                ApplyToControl(child, p);
        }

        private static void EnsureNativeScrollTheme(Control control)
        {
            ApplyNativeScrollTheme(control);
            control.HandleCreated -= NativeScrollThemeHandleCreated;
            control.HandleCreated += NativeScrollThemeHandleCreated;
        }

        private static void NativeScrollThemeHandleCreated(object? sender, EventArgs e)
        {
            if (sender is Control control)
                ApplyNativeScrollTheme(control);
        }

        private static void ApplyNativeScrollTheme(Control control)
        {
            if (!control.IsHandleCreated) return;
            bool dark = control.BackColor.GetBrightness() < 0.5f;
            try { SetWindowTheme(control.Handle, dark ? "DarkMode_Explorer" : "Explorer", null); }
            catch { /* best effort only */ }
        }

        private static void ComboDrawItem(object? sender, DrawItemEventArgs e)
        {
            if (sender is not ComboBox combo || e.Index < 0)
                return;

            var p = combo.Tag as ThemePalette ?? ThemePalette.Light();
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            using var back = new SolidBrush(selected ? ControlPaint.Light(p.Surface, 0.12f) : p.Back);
            using var fore = new SolidBrush(p.Text);

            e.Graphics.FillRectangle(back, e.Bounds);
            var text = combo.GetItemText(combo.Items[e.Index]);
            TextRenderer.DrawText(e.Graphics, text, combo.Font, e.Bounds, p.Text, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            e.DrawFocusRectangle();
        }

        private static void TabDrawItem(object? sender, DrawItemEventArgs e)
        {
            if (sender is not TabControl tab || e.Index < 0)
                return;

            var p = tab.Tag as ThemePalette ?? ThemePalette.Light();
            bool selected = tab.SelectedIndex == e.Index;
            var bounds = tab.GetTabRect(e.Index);
            using var back = new SolidBrush(selected ? p.Back : p.Surface);
            using var border = new Pen(p.Border);

            e.Graphics.FillRectangle(back, bounds);
            e.Graphics.DrawRectangle(border, bounds);
            TextRenderer.DrawText(
                e.Graphics,
                tab.TabPages[e.Index].Text,
                tab.Font,
                bounds,
                p.Text,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hWnd, string? pszSubAppName, string? pszSubIdList);
    }
}
