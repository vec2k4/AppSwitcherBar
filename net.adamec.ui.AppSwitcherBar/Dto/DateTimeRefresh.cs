using net.adamec.ui.AppSwitcherBar.ViewModel;
using System;
using System.Threading.Tasks;

namespace net.adamec.ui.AppSwitcherBar.Dto
{
    internal class DateTimeRefresh
    {
        private static bool isRefreshStopped = false;
        private static Task? refreshTask = null;

        private DateTimeRefresh()
        {
        }

        public static void StopRefreshing()
        {
            isRefreshStopped = true;
        }

        public static void StartRefreshing(MainViewModel vm)
        {
            if (refreshTask != null)
                return;

            refreshTask = Task.Factory.StartNew(async () =>
            {
                while (!isRefreshStopped)
                {
                    var dt = DateTime.Now;
                    vm.CurrentDate = $"{dt:dd.MM.yyyy}";
                    vm.CurrentTime = $"{dt:HH:mm:ss}";
                    await Task.Delay(250);
                }

                refreshTask = null;
            }, TaskCreationOptions.LongRunning);
        }
    }
}
