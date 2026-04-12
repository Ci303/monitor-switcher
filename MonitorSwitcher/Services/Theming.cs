using System;
using System.Drawing;
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
                    lbl.BackColor = Color.Transparent;
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

                case Panel pnl:
                    // Let 1px rule lines keep their custom color; other panels follow background
                    if (pnl.Height > 2)
                        pnl.BackColor = p.Back;
                    break;

                case DataGridView dgv:
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

                    dgv.RowHeadersDefaultCellStyle.BackColor = p.Surface;
                    dgv.RowHeadersDefaultCellStyle.ForeColor = p.Text;
                    dgv.RowHeadersDefaultCellStyle.SelectionBackColor = p.Surface;
                    dgv.RowHeadersDefaultCellStyle.SelectionForeColor = p.Text;
                    break;

                default:
                    c.BackColor = p.Back;
                    c.ForeColor = p.Text;
                    break;
            }

            foreach (Control child in c.Controls)
                ApplyToControl(child, p);
        }
    }
}
