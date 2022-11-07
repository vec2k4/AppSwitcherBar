using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace net.adamec.ui.AppSwitcherBar.Win32.Services
{
    internal class Taskbar
    {
        [DllImport("user32.dll")]
        private static extern int FindWindow(string className, string windowText);

        [DllImport("user32.dll")]
        private static extern int ShowWindow(int hwnd, int command);

        [DllImport("user32.dll")]
        public static extern int FindWindowEx(int parentHandle, int childAfter, string className, int windowTitle);

        [DllImport("user32.dll")]
        private static extern int GetDesktopWindow();

        private const int SW_HIDE = 0;
        private const int SW_SHOW = 1;

        private static bool isVisible = true;
        private static bool isRefreshStopped = false;
        private static Task? visibilityHandlingTask = null;

        protected static int Handle
        {
            get
            {
                return FindWindow("Shell_TrayWnd", "");
            }
        }

        protected static int HandleOfStartButton
        {
            get
            {
                int handleOfDesktop = GetDesktopWindow();
                int handleOfStartButton = FindWindowEx(handleOfDesktop, 0, "button", 0);
                return handleOfStartButton;
            }
        }

        private Taskbar()
        {
            // hide ctor
        }

        public static void Show()
        {
            isVisible = true;
            RefreshTaskbarVisibility();
        }

        public static void Hide()
        {
            isVisible = false;
            RefreshTaskbarVisibility();
        }

        public static void StopTaskbarVisibilityRefresh()
        {
            isRefreshStopped = true;
        }

        public static void RefreshTaskbarVisibility()
        {
            if (visibilityHandlingTask != null)
                return;

            visibilityHandlingTask = Task.Factory.StartNew(async () =>
            {
                while (!isRefreshStopped)
                {
                    if (isVisible)
                    {
                        _ = ShowWindow(Handle, SW_SHOW);
                        _ = ShowWindow(HandleOfStartButton, SW_SHOW);
                    }
                    else
                    {
                        _ = ShowWindow(Handle, SW_HIDE);
                        _ = ShowWindow(HandleOfStartButton, SW_HIDE);
                    }
                    await Task.Delay(250);
                }

                _ = ShowWindow(Handle, SW_SHOW);
                _ = ShowWindow(HandleOfStartButton, SW_SHOW);
                visibilityHandlingTask = null;
            }, TaskCreationOptions.LongRunning);
        }
    }
}

