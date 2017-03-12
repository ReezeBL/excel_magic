using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Meister_Hämmerlein.WindowForms;

namespace Meister_Hämmerlein
{
    internal class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        public static void ShowConsoleWindow()
        {
            var handle = GetConsoleWindow();

            if (handle == IntPtr.Zero)
            {
                AllocConsole();
            }
            else
            {
                ShowWindow(handle, SW_SHOW);
            }
        }

        public static void HideConsoleWindow()
        {
            var handle = GetConsoleWindow();

            ShowWindow(handle, SW_HIDE);
        }


        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;

        [STAThread]
        private static void Main(string[] args)
        {
            if(args.Length > 0)
                ShowConsoleWindow();
            Application.Run(new MainWindow());
        }
    }
}
