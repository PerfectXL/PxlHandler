using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace PxlHandler
{
    internal class ProcessHelper
    {
        public static void BringProcessToFront(string processName)
        {
            try
            {
                Process process = Process.GetProcessesByName(processName).FirstOrDefault();
                if (process == null)
                {
                    return;
                }

                process.WaitForInputIdle();
                IntPtr s = process.MainWindowHandle;
                SetForegroundWindow(s);

                Console.WriteLine($"Process found: {process.ProcessName} ({process.Id})");
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Error in {nameof(BringProcessToFront)}. {exception.Message}");
            }
        }

        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}