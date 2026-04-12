#nullable enable
using System;
using System.Drawing;
using System.Windows.Forms;
using WorkMonitorSwitcher.Services; // Themer + DwmInterop

namespace WorkMonitorSwitcher.UI
{
    public static class ThemedMessageBox
    {
        public static DialogResult Info(IWin32Window owner, string text, string title, bool dark)
            => Show(owner, text, title, dark, SystemIcons.Information);

        public static DialogResult Warn(IWin32Window owner, string text, string title, bool dark)
            => Show(owner, text, title, dark, SystemIcons.Warning);

        public static DialogResult Error(IWin32Window owner, string text, string title, bool dark)
            => Show(owner, text, title, dark, SystemIcons.Error);

        private static DialogResult Show(IWin32Window owner, string text, string title, bool dark, Icon icon)
        {
            // Dialog shell
            using var dlg = new Form
            {
                Text = title,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = false,
                AutoScaleMode = AutoScaleMode.Dpi,
                Width = 200,
                Height = 170
            };

            // Icon
            var pb = new PictureBox
            {
                Image = icon.ToBitmap(),
                SizeMode = PictureBoxSizeMode.CenterImage,
                Location = new Point(16, 20),
                Size = new Size(32, 32)
            };

            // Text
            var lbl = new Label
            {
                AutoSize = false,
                Location = new Point(60, 18),
                Size = new Size(dlg.ClientSize.Width - 76, 70),
                Text = text
            };

            // OK button
            var ok = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Size = new Size(80, 28),
            };
            ok.Location = new Point(dlg.ClientSize.Width - ok.Width - 16,
                                    dlg.ClientSize.Height - ok.Height - 16);
            ok.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            dlg.AcceptButton = ok;
            dlg.Controls.Add(pb);
            dlg.Controls.Add(lbl);
            dlg.Controls.Add(ok);

            // Apply theme
            var palette = dark ? ThemePalette.Dark() : ThemePalette.Light();
            Themer.Apply(dlg, palette);
            try { DwmInterop.SetDarkTitleBar(dlg.Handle, dark); } catch { /* ignore */ }

            return dlg.ShowDialog(owner);
        }
    }
}
