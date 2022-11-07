using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using net.adamec.ui.AppSwitcherBar.Config;
using net.adamec.ui.AppSwitcherBar.Dto;
using net.adamec.ui.AppSwitcherBar.ViewModel;
using net.adamec.ui.AppSwitcherBar.Win32.Services;
using System.ComponentModel;

namespace net.adamec.ui.AppSwitcherBar
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {

        public MainWindow(IOptions<AppSettings> options, ILogger<MainWindow> logger) : base(options, logger)
        {
            InitializeComponent();

            if (DataContext is MainViewModel viewModel && !IsDesignTime)
            {
                //initialize the "active" logic of view model - retrieving the information about windows
                Loaded += (_, _) =>
                {
                    Taskbar.Show();
                    viewModel.Init(Hwnd);
                };
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            Taskbar.StopTaskbarVisibilityRefresh();
            DateTimeRefresh.StopRefreshing();
            base.OnClosing(e);
        }
    }
}
