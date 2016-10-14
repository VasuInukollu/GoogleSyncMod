using System;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    static class Program
    {
        private const string MUTEXGUID = "ACBBBC09-F76C-4874-AAFF-4F3353A5A5A6";
        private static Mutex mutex;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");
            //prevent more than one instance of the program
            if (IsRunning())
            {   //Instance already exists, so show only Main-Window  
                NativeMethods.PostMessage((IntPtr)NativeMethods.HWND_BROADCAST, NativeMethods.WM_GCSM_SHOWME, IntPtr.Zero, IntPtr.Zero);
                return;
            }
            else
            {
                RegisterEventHandlers();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(SettingsForm.Instance);
            }
            GC.KeepAlive(mutex);
        }

        private static void RegisterEventHandlers()
        {
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
        }

        public static bool IsRunning()
        {
            bool ok;
            mutex = new Mutex(true, MUTEXGUID, out ok);
            if (mutex.WaitOne(TimeSpan.Zero, true))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Fallback. If there is some try/catch missing we will handle it here, just before the application quits unhandled
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception)
                ErrorHandler.Handle((Exception)e.ExceptionObject);
        }
    }
}