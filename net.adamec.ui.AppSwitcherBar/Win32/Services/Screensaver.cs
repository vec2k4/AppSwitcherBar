using System;
using System.Runtime.InteropServices;

namespace net.adamec.ui.AppSwitcherBar.Win32.Services
{
    internal class Screensaver
    {
        [Flags]
        public enum EXECUTION_STATE : uint
        {
            ES_AWAYMODE_REQUIRED = 0x00000040,
            ES_CONTINUOUS = 0x80000000,
            ES_DISPLAY_REQUIRED = 0x00000002,
            ES_SYSTEM_REQUIRED = 0x00000001
            // Legacy flag, should not be used.
            // ES_USER_PRESENT = 0x00000004
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern EXECUTION_STATE SetThreadExecutionState(EXECUTION_STATE esFlags);

        private Screensaver()
        {
        }

        public static bool IsDisabled
        {
            get; private set;
        }

        public static void EnableScreensaver()
        {
            SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS);
            IsDisabled = false;
        }

        public static void DisableScreensaver()
        {
            SetThreadExecutionState(EXECUTION_STATE.ES_DISPLAY_REQUIRED | EXECUTION_STATE.ES_CONTINUOUS);
            IsDisabled = true;
        }
    }
}
