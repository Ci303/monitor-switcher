using System;
using System.Windows.Forms;

namespace WorkMonitorSwitcher
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize(); // <-- Keep this if it's already part of your project

            Application.Run(new Form1());
        }
    }
}