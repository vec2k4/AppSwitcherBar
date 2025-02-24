﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using MahApps.Metro.IconPacks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using net.adamec.ui.AppSwitcherBar.AppBar;
using net.adamec.ui.AppSwitcherBar.Config;
using net.adamec.ui.AppSwitcherBar.Dto;
using net.adamec.ui.AppSwitcherBar.Dto.Search;
using net.adamec.ui.AppSwitcherBar.Win32.NativeInterfaces.Extensions;
using net.adamec.ui.AppSwitcherBar.Win32.Services;
using net.adamec.ui.AppSwitcherBar.Win32.Services.JumpLists;
using net.adamec.ui.AppSwitcherBar.Win32.Services.Shell;
using net.adamec.ui.AppSwitcherBar.Win32.Services.Shell.Properties;
using net.adamec.ui.AppSwitcherBar.Win32.Services.Startup;
using net.adamec.ui.AppSwitcherBar.Wpf;
using static net.adamec.ui.AppSwitcherBar.Dto.PinnedAppInfo;

// ReSharper disable StringLiteralTypo
// ReSharper disable IdentifierTypo
// ReSharper disable CommentTypo

namespace net.adamec.ui.AppSwitcherBar.ViewModel
{
    /// <summary>
    /// The ViewModel for <see cref="MainWindow"/>.
    /// Encapsulates the data and logic related to "task bar applications/windows"
    ///  - pulling the list of them, switching the apps, presenting the thumbnails
    /// </summary>
    public partial class MainViewModel : INotifyPropertyChanged
    {
        /// <summary>
        /// Native handle (HWND) of the <see cref="MainWindow"/>
        /// </summary>
        private IntPtr mainWindowHwnd;
        /// <summary>
        /// Native handle of the application window thumbnail currently shown
        /// </summary>
        private IntPtr thumbnailHandle;

        /// <summary>
        /// <see cref="DispatcherTimer"/> used to periodically pull (refresh) the information about (open) application windows
        /// </summary>
        private readonly DispatcherTimer timer;

        /// <summary>
        /// <see cref="BackgroundWorker"/> used to retrieve helper data on background
        /// </summary>
        private readonly BackgroundWorker backgroundInitWorker;

        /// <summary>
        /// Flag whether the background data have been retrieved
        /// </summary>
        private bool backgroundDataRetrieved;

        /// <summary>
        /// Flag whether the background data have been retrieved
        /// </summary>
        public bool BackgroundDataRetrieved
        {
            get => backgroundDataRetrieved;
            set
            {
                if (backgroundDataRetrieved != value)
                {
                    backgroundDataRetrieved = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Flag whether the <see cref="RefreshAllWindowsCollection"/> method is run for the first time
        /// </summary>
        private bool isFirstRun = true;


        /// <summary>
        /// Native handle of the last known foreground window
        /// </summary>
        private IntPtr lastForegroundWindow = IntPtr.Zero;

        /// <summary>
        /// Information about the applications installed in system
        /// </summary>
        private readonly InstalledApplications installedApplications = new();

        /// <summary>
        /// Information about the applications pinned in the taskbar
        /// </summary>
        private PinnedAppInfo[] pinnedApplications = Array.Empty<PinnedAppInfo>();

        /// <summary>
        /// Information about the known folder paths and GUIDs
        /// </summary>
        private StringGuidPair[] knownFolders = Array.Empty<StringGuidPair>();

        /// <summary>
        /// Dictionary of known AppIds from configuration containing pairs executable-appId (the key is in lower case)
        /// When built from configuration, the record (key) is created for full path from config and another one without a path (file name only) if applicable
        /// </summary>
        private readonly Dictionary<string, string> knownAppIds;

        /// <summary>
        /// Application settings
        /// </summary>
        public IAppSettings Settings { get; }




        /// <summary>
        /// Flag whether the option to set Run On Windows Startup is available
        /// </summary>
        public bool RunOnWinStartupAvailable => Settings.AllowRunOnWindowsStartup;

        /// <summary>
        /// Flag whether the AppSwitcherBar is set to run on Windows startup
        /// </summary>
        private bool runOnWinStartupSet;
        /// <summary>
        /// Flag whether the AppSwitcherBar is set to run on Windows startup
        /// </summary>
        public bool RunOnWinStartupSet
        {
            get => runOnWinStartupSet;
            set
            {
                if (runOnWinStartupSet != value)
                {
                    runOnWinStartupSet = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool isWindowsTaskbarHidden;

        public bool IsWindowsTaskbarHidden
        {
            get => isWindowsTaskbarHidden;
            set
            {
                if (isWindowsTaskbarHidden != value)
                {
                    isWindowsTaskbarHidden = value;
                    if (isWindowsTaskbarHidden)
                    {
                        Taskbar.Hide();
                    }
                    else
                    {
                        Taskbar.Show();
                        Application.Current.MainWindow.Focus();
                    }
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(TaskbarToggleButtonTooltip));
                }
            }
        }

        public string TaskbarToggleButtonTooltip => isWindowsTaskbarHidden ? "Taskbar hidden" : "Taskbar visible";


        private bool isScreensaverDisabled;

        public bool IsScreensaverDisabled
        {
            get => isScreensaverDisabled;
            set
            {
                if (isScreensaverDisabled != value)
                {
                    isScreensaverDisabled = value;
                    if (isScreensaverDisabled)
                    {
                        Screensaver.DisableScreensaver();
                    }
                    else
                    {
                        Screensaver.EnableScreensaver();
                    }
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(ScreensaverToggleButtonTooltip));
                }
            }
        }

        public string ScreensaverToggleButtonTooltip => isScreensaverDisabled ? "Screensaver disabled" : "Screensaver enabled";

        private string currentDate = "dd.MM.YYYY";

        public string CurrentDate
        {
            get => currentDate;
            set
            {
                if (!currentDate.Equals(value))
                {
                    currentDate = value;
                    OnPropertyChanged();
                }
            }
        }

        private string currentTime = "HH:mm:ss";

        public string CurrentTime
        {
            get => currentTime;
            set
            {
                if (!currentTime.Equals(value))
                {
                    currentTime = value;
                    OnPropertyChanged();
                }
            }
        }


        /// <summary>
        /// Flag whether the search is in progress
        /// </summary>
        private bool isInSearch;

        /// <summary>
        /// Flag whether the search is in progress
        /// </summary>
        public bool IsInSearch
        {
            get => isInSearch;
            set
            {
                if (isInSearch != value)
                {
                    isInSearch = value;
                    if (isInSearch)
                    {
                        InitSearch();
                    }
                    else
                    {
                        EndSearch();
                    }
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Flag whether the search has results
        /// </summary>
        private bool hasSearchResults;

        /// <summary>
        /// Flag whether the search has results
        /// </summary>
        public bool HasSearchResults
        {
            get => hasSearchResults;
            private set
            {
                if (hasSearchResults != value)
                {
                    hasSearchResults = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Text to be searched for
        /// </summary>
        private string? searchText;

        /// <summary>
        /// Text to be searched for
        /// </summary>
        public string? SearchText
        {
            get => searchText;
            set
            {
                if (searchText != value)
                {
                    searchText = value;
                    DoSearch(searchText);
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Array of the screen edges the app-bar can be docked to
        /// </summary>
        public AppBarDockMode[] Edges { get; } = {
            AppBarDockMode.Left,
            AppBarDockMode.Right,
            AppBarDockMode.Top,
            AppBarDockMode.Bottom};

        /// <summary>
        /// Array of the information about all monitors (displays)
        /// </summary>
        public MonitorInfo[] AllMonitors { get; }

        /// <summary>
        /// Application window buttons group manager
        /// </summary>
        public AppButtonManager ButtonManager { get; }

        /// <summary>
        /// Command requesting an "ad-hoc" refresh of the list of application windows (no param used)
        /// </summary>
        public ICommand RefreshWindowCollectionCommand { get; }
        /// <summary>
        /// Command requesting the toggle of the application window.
        /// Switch the application window to foreground or minimize it
        /// The command parameter is HWND of the application window
        /// </summary>
        public ICommand ToggleApplicationWindowCommand { get; }
        /// <summary>
        /// Command requesting the render of application window thumbnail into the popup
        /// The command parameter is fully populated <see cref="ThumbnailPopupCommandParams"/> object
        /// </summary>
        public ICommand ShowThumbnailCommand { get; }
        /// <summary>
        /// Command requesting to hide application window thumbnail (no param used)
        /// </summary>
        public ICommand HideThumbnailCommand { get; }

        /// <summary>
        /// Command requesting to build the context menu for application window button
        /// </summary>
        public ICommand BuildContextMenuCommand { get; }

        /// <summary>
        /// Command requesting to toggle Run on Windows startup - set/remove the startup link
        /// </summary>
        public ICommand ToggleRunOnStartupCommand { get; }

        /// <summary>
        /// Command requesting to launch pinned application
        /// </summary>
        public ICommand LaunchPinnedAppCommand { get; }


        /// <summary>
        /// Command sending a special key press related to search
        /// </summary>
        public ICommand SearchSpecialKeyCommand { get; }

        /// <summary>
        /// JumpList service to be used
        /// </summary>
        private readonly IJumpListService jumpListService;

        /// <summary>
        /// Startup service to be used
        /// </summary>
        private readonly IStartupService startupService;

        /// <summary>
        /// Name of the Feature Flag for windows anonymization
        /// </summary>
        // ReSharper disable once InconsistentNaming
        private const string FF_AnonymizeWindows = "AnonymizeWindows";

        /// <summary>
        /// Name of the Feature Flag for using the undocumented application resolver to get the app it
        /// </summary>
        // ReSharper disable once InconsistentNaming
        private const string FF_UseApplicationResolver = "UseApplicationResolver";

        /// <summary>
        /// Map used for simple anonymization
        /// </summary>
        ///
        private readonly Dictionary<char, char> anonymizeMap = new();

        /// <summary>
        /// Internal CTOR
        /// Directly used by <see cref="ViewModelLocator"/> when creating a design time instance.
        /// Internally called by public "DI bound" CTOR
        /// </summary>
        /// <param name="settings">Application setting</param>
        /// <param name="logger">Logger to be used</param>
        /// <param name="jumpListService">JumpList service to be used</param>
        /// <param name="startupService">Startup service to be used</param>
        internal MainViewModel(IAppSettings settings, ILogger logger, IJumpListService jumpListService, IStartupService startupService)
        {
            this.logger = logger;
            this.jumpListService = jumpListService;
            this.startupService = startupService;

            Settings = settings;
            AllMonitors = Monitor.GetAllMonitors();
            ButtonManager = new AppButtonManager(Settings);

            RefreshWindowCollectionCommand = new RelayCommand(RefreshAllWindowsCollection);
            ToggleApplicationWindowCommand = new RelayCommand(ToggleApplicationWindow);
            ShowThumbnailCommand = new RelayCommand(ShowThumbnail);
            HideThumbnailCommand = new RelayCommand(HideThumbnail);
            BuildContextMenuCommand = new RelayCommand(BuildContextMenu);
            ToggleRunOnStartupCommand = new RelayCommand(ToggleRunOnWinStartup);
            LaunchPinnedAppCommand = new RelayCommand(LaunchPinnedApp);
            SearchSpecialKeyCommand = new RelayCommand(SearchSpecialKey);

            runOnWinStartupSet = startupService.HasAppStartupLink();

            knownAppIds = settings.GetKnowAppIds();

            if (settings.FeatureFlag(FF_AnonymizeWindows, false))
            {
                InitAnonymizeMap();
            }

            timer = new DispatcherTimer(DispatcherPriority.Render)
            {
                Interval = TimeSpan.FromMilliseconds(Settings.RefreshWindowInfosIntervalMs)
            };
            timer.Tick += (_, _) => { if (!ButtonManager.IsBusy) RefreshAllWindowsCollection(false); };

            backgroundInitWorker = new BackgroundWorker();
            backgroundInitWorker.DoWork += (_, eventArgs) => { eventArgs.Result = RetrieveBackgroundData(); };
            backgroundInitWorker.RunWorkerCompleted += (_, eventArgs) => { OnBackgroundDataRetrieved((BackgroundData)eventArgs.Result!); };
            this.startupService = startupService;
        }


        /// <summary>
        /// CTOR used by DI
        /// </summary>
        /// <param name="options">Application settings configuration</param>
        /// <param name="logger">Logger to be used</param>
        /// <param name="jumpListService">JumpList service to be used</param>
        /// <param name="startupService">Startup service to be used</param>
        // ReSharper disable once UnusedMember.Global
        public MainViewModel(IOptions<AppSettings> options, ILogger<MainViewModel> logger, IJumpListService jumpListService, IStartupService startupService)
            : this(options.Value, logger, jumpListService, startupService)
        {
            //used from DI - DI populates the parameters and the internal CTOR is called then
        }

        /// <summary>
        /// (Late) initialize the view model.
        /// Registers the native handle (HWND) of the <see cref="MainWindow"/> and
        /// starts the <see cref="timer"/> used to periodically populate the information about application windows
        /// </summary>
        /// <param name="mainWndHwnd">Native handle (HWND) of the <see cref="MainWindow"/></param>
        public void Init(IntPtr mainWndHwnd)
        {
            mainWindowHwnd = mainWndHwnd;

            timer.Start();


        }

        /// <summary>
        /// Initialize character map for simple anonymization
        /// </summary>
        private void InitAnonymizeMap()
        {
            var rnd = new Random(DateTime.Now.GetHashCode());
            for (var c = 'a'; c <= 'z'; c++)
            {
                anonymizeMap[c] = (char)('a' + rnd.Next(26));
            }
            for (var c = 'A'; c <= 'Z'; c++)
            {
                anonymizeMap[c] = (char)('A' + rnd.Next(26));
            }
            for (var c = '0'; c <= '9'; c++)
            {
                anonymizeMap[c] = (char)('0' + rnd.Next(26));
            }
            foreach (var c in " -.:/")
            {
                anonymizeMap[c] = (char)('a' + rnd.Next(26));
            }
        }

        /// <summary>
        /// Simple anonymize given string
        /// </summary>
        /// <param name="s">string to anonymize</param>
        /// <param name="saltInt">anonymization salt</param>
        /// <returns>anonymized string</returns>
        private string? Anonymize(string? s, int saltInt)
        {
            if (s == null) return null;
            var salt = saltInt.ToString();
            while (salt.Length < s.Length)
            {
                salt += salt;
            }

            var retVal = string.Empty;
            for (var i = 0; i < s.Length; i++)
            {
                var c = s[i];
                var cs = (char)(c + salt[i]);
                retVal += anonymizeMap.TryGetValue(cs, out var a) ? a : anonymizeMap.TryGetValue(c, out a) ? a : c;
            }

            return retVal;
        }

        /// <summary>
        /// Retrieve helper data on background
        /// </summary>
        internal BackgroundData? RetrieveBackgroundData()
        {
            var timestampStart = DateTime.Now;
            var timestampEndInstalledApps = DateTime.MinValue;

            bool isSuccess;
            string resultMsg;
            BackgroundData? data = null;

            try
            {
                //retrieve installed apps -> AUMIs, icons
                var dataInstalledApps = new List<InstalledApplication>();
                var appsFolder = Shell.GetAppsFolder();
                if (appsFolder != null)
                {
                    Shell.EnumShellItems(appsFolder, item =>
                     {
                         var appName = item.GetDisplayName();
                         var propertyStore = item.GetPropertyStore();
                         var shellProperties = propertyStore?.GetProperties() ?? new ShellPropertiesSubset(); //empty object
                         if (!shellProperties.IsApplication) return; //not application (not runnable)

                         var appUserModelId = shellProperties.ApplicationUserModelId;
                         var iconSource = Shell.GetShellItemBitmapSource(item, 32);
                         var lnkTarget = propertyStore?.GetPropertyValue<string>(PropertyKey.PKEY_Link_TargetParsingPath);

                         dataInstalledApps.Add(new InstalledApplication(appName, appUserModelId, lnkTarget, iconSource, shellProperties));
                         LogInstalledAppInfo(appName, appUserModelId ?? "[Unknown]", iconSource != null, lnkTarget ?? "[N/A]");
                     });
                }

                timestampEndInstalledApps = DateTime.Now;

                data = new BackgroundData(dataInstalledApps.ToArray());
                resultMsg = "OK";
                isSuccess = true;
            }
            catch (Exception ex)
            {
                isSuccess = false;
                resultMsg = $"{ex.GetType().Name}: {ex.Message}";
            }

            var timestampEndTotal = DateTime.Now;
            var durationInstalledApps = (timestampEndInstalledApps - timestampStart).TotalMilliseconds;
            var durationTotal = (timestampEndTotal - timestampStart).TotalMilliseconds;

            LogBackgroundDataInitTelemetry(
                isSuccess ? LogLevel.Information : LogLevel.Error,
                DateTime.Now, isSuccess, resultMsg,
                (int)durationTotal, (int)durationInstalledApps);


            return data;
        }

        /// <summary>
        /// Called when the background data have been retrieved - "copy" them to view model
        /// </summary>
        /// <param name="data">Retrieved background data or null when not available</param>
        internal void OnBackgroundDataRetrieved(BackgroundData? data)
        {
            if (data != null)
            {
                installedApplications.Clear();
                foreach (var installedApplication in data.InstalledApplications)
                {
                    if (installedApplication.IconSource != null)
                    {
                        //clone the object, co it can be uses in other (UI) thread!!!
                        installedApplication.IconSource = installedApplication.IconSource.Clone();
                    }
                    installedApplications.Add(installedApplication);
                }
            }
            BackgroundDataRetrieved = true;
        }

        /// <summary>
        /// Pulls the information about available application windows and updates <see cref="ButtonManager"/> window collection.
        /// </summary>
        /// <param name="hardRefresh">When the parameter is bool and true, it forces the hard refresh.
        /// The <see cref="ButtonManager"/>window collection is cleared first and the background data are refreshed on hard refresh.
        ///  Otherwise just the window collection is updated</param>
        private void RefreshAllWindowsCollection(object? hardRefresh)
        {
            var isHardRefresh = (hardRefresh is bool b || bool.TryParse(hardRefresh?.ToString(), out b)) && b;

            if (isFirstRun)
            {
                isHardRefresh = true;
                isFirstRun = false;
            }

            if (isHardRefresh)
            {
                //get known folders
                knownFolders = Shell.GetKnownFolders();
                //get information about pinned applications
                pinnedApplications = jumpListService.GetPinnedApplications(knownFolders);

                ButtonManager.BeginHardRefresh(pinnedApplications);

                //Refresh also init data
                if (!backgroundInitWorker.IsBusy)
                {
                    BackgroundDataRetrieved = false;
                    backgroundInitWorker.RunWorkerAsync();
                }
            }
            else
            {
                //begin update of windows collection
                ButtonManager.BeginUpdate();
            }

            //Retrieve the current foreground window
            var foregroundWindow = WndAndApp.GetForegroundWindow();
            if (foregroundWindow != mainWindowHwnd)
            {
                lastForegroundWindow = foregroundWindow; //"filter out" the main window as being the foreground one to proper handle the toggle

                if (IsInSearch) EndSearch(); //cancel search when other app get's focus
            }

            //Enum windows
            WndAndApp.EnumVisibleWindows(
                mainWindowHwnd,
                hwnd => ButtonManager[hwnd],
                (hwnd, wnd, caption, threadId, processId, ptrProcess) =>
                {
                    //caption anonymization
                    if (Settings.FeatureFlag<bool>(FF_AnonymizeWindows))
                    {
                        var appName =
                            installedApplications.GetInstalledApplicationFromAppId(wnd?.AppId ?? string.Empty)?.Name ??
                            installedApplications.GetInstalledApplicationFromExecutable(wnd?.Executable ?? string.Empty)?.Name;

                        if (caption != appName)
                        {
                            caption = (string.IsNullOrEmpty(appName) ? "" : $"{appName} - ") + Anonymize(caption, hwnd.ToInt32());
                        }
                    }


                    //Check whether it's a "new" application window or a one already existing in the ButtonManager
                    if (wnd == null)
                    {
                        //app executable
                        var executable = WndAndApp.GetProcessExecutable(ptrProcess);
                        //new window
                        wnd = new WndInfo(hwnd, caption, threadId, processId, executable);

                        //Try to get AppUserModelId using the win32 app resolver, no need to wait for background data
                        if (Settings.FeatureFlag<bool>(FF_UseApplicationResolver))
                            wnd.AppId = WndAndApp.GetWindowApplicationUserModelId(hwnd);
                    }
                    else
                    {
                        //existing (known) window
                        wnd.MarkToKeep(); //reset the "remove from collection" flag
                        wnd.Title = caption; //update the title
                    }

                    wnd.IsForeground = hwnd == lastForegroundWindow; //check whether the window is foreground window (will be highlighted in UI)


                    if ((wnd.AppId is null || Settings.CheckForAppIdChange) && BackgroundDataRetrieved) // || isHardRefresh - not needed as wnd will be new with AppId =null
                    {
                        string? appUserModelId = null;

                        //Try to get AppUserModelId using the win32 app resolver
                        if (Settings.FeatureFlag<bool>(FF_UseApplicationResolver))
                        {
                            appUserModelId = WndAndApp.GetWindowApplicationUserModelId(hwnd);
                        }

                        //Try to get AppUserModelId from window - for windows that explicitly define the AppId
                        if (appUserModelId == null)
                        {
                            var store = Shell.GetPropertyStoreForWindow(wnd.Hwnd);
                            if (store != null)
                            {
                                var hr = store.GetCount(out var c);
                                if (hr.IsSuccess && c > 0)
                                {
                                    //try to get AppUserModelId property
                                    appUserModelId = store.GetPropertyValue<string>(PropertyKey.PKEY_AppUserModel_ID);
                                    var shellProperties = store.GetProperties();
                                    wnd.ShellProperties = shellProperties;
                                }
                            }
                        }

                        //Try the app ids from configuration
                        //It must contain the record for (shell) explorer (done in CTOR) as it will not work properly for explorer windows without this hack
                        if (appUserModelId == null && wnd.Executable != null &&
                            (knownAppIds.TryGetValue(wnd.Executable.ToLowerInvariant(), out var appId) ||
                             knownAppIds.TryGetValue(Path.GetFileName(wnd.Executable.ToLowerInvariant()), out appId)))
                        {
                            appUserModelId = appId;
                        }


                        //try to get AppUserModelId from process if not "at window"
                        appUserModelId ??= WndAndApp.GetProcessApplicationUserModelId(ptrProcess);

                        if (appUserModelId == null && !string.IsNullOrEmpty(wnd.Executable))
                        {
                            //try to get from installed app (identified by executable) or use executable as fallback
                            appUserModelId = installedApplications.GetAppIdFromExecutable(wnd.Executable, out var _);
                        }

                        wnd.AppId = appUserModelId;
                    }



                    if (Settings.CheckForIconChange || isHardRefresh)
                    {
                        //Try to retrieve the window icon
                        wnd.BitmapSource = WndAndApp.GetWindowIcon(hwnd);

                        if (wnd.BitmapSource == null)
                        {
                            //try to get icon from installed application
                            if (!string.IsNullOrEmpty(wnd.AppId))
                            {
                                wnd.BitmapSource = installedApplications.GetInstalledApplicationFromAppId(wnd.AppId)?.IconSource;
                            }
                        }

                        if (Settings.InvertWhiteIcons)
                            wnd.BitmapSource = Resource.InvertBitmapIfWhiteOnly(wnd.BitmapSource);
                    }

                    //Add new window to button manager and the button windows collections
                    if (wnd.ChangeStatus == WndInfo.ChangeStatusEnum.New)
                    {
                        ButtonManager.Add(wnd);
                    }

                    if (isHardRefresh)
                    {
                        LogEnumeratedWindowInfo(wnd.ToString());
                    }

                }); //enum visible windows

            ButtonManager.EndUpdate();

        }

        /// <summary>
        /// Switch the application window with given <paramref name="hwnd"/> to foreground or minimize it.
        /// </summary>
        /// <remarks>
        /// The function doesn't throw any exception when the handle is invalid, it just ignores it end "silently" returns
        /// </remarks>
        /// <param name="hwnd">Native handle (HWND) of the application window</param>
        private void ToggleApplicationWindow(object? hwnd)
        {
            if (hwnd is not IntPtr hWndIntPtr || hWndIntPtr == IntPtr.Zero) return; //invalid command parameter, do nothing
            ToggleApplicationWindow(hWndIntPtr, false);
        }

        /// <summary>
        /// Switch the application window with given <paramref name="hwnd"/> to foreground or minimize it (if <paramref name="forceActivate"/> is not set).
        /// </summary>
        /// <remarks>
        /// The function doesn't throw any exception when the handle is invalid, it just ignores it end "silently" returns
        /// </remarks>
        /// <param name="hwnd">Native handle (HWND) of the application window</param>
        /// <param name="forceActivate">When the flag is set, the window is always activated. When it's false and the window is foreground already, it's minimized</param>
        private void ToggleApplicationWindow(IntPtr hwnd, bool forceActivate)
        {
            if (hwnd == IntPtr.Zero) return; //invalid command parameter, do nothing

            //got the handle, get the window information
            var wnd = ButtonManager[hwnd];
            if (wnd is null) return; //unknown window, do nothing

            if (wnd.IsForeground && !forceActivate)
            {
                //it's a foreground window - minimize it and return
                WndAndApp.MinimizeWindow(hwnd);
                LogMinimizeApp(hwnd, wnd.Title);
                return;
            }

            var wasMinimized = WndAndApp.ActivateWindow(hwnd);
            LogSwitchApp(hwnd, wnd.Title, wasMinimized);

            //refresh the window list
            RefreshAllWindowsCollection(false);
        }

        /// <summary>
        /// Registers and shows the application window thumbnail within the popup
        /// Parameter <paramref name="param"/> must be <see cref="ThumbnailPopupCommandParams"/> object,
        /// encapsulating <see cref="ThumbnailPopupCommandParams.SourceHwnd"/> of application window,
        /// <see cref="ThumbnailPopupCommandParams.TargetHwnd"/> of the popup window and
        /// <see cref="ThumbnailPopupCommandParams.TargetRect"/> with the bounding box within the popup window.
        /// </summary>
        /// <param name="param"><see cref="ThumbnailPopupCommandParams"/> object with source and target information</param>
        /// <exception cref="ArgumentException">When the <paramref name="param"/> is not <see cref="ThumbnailPopupCommandParams"/> object or is null, <see cref="ArgumentException"/> is thrown</exception>
        private void ShowThumbnail(object? param)
        {
            if (param is not ThumbnailPopupCommandParams cmdParams)
            {
                LogWrongCommandParameter(nameof(ThumbnailPopupCommandParams));
                throw new ArgumentException($"Command parameter must be {nameof(ThumbnailPopupCommandParams)}", nameof(param));
            }

            HideThumbnail(); //unregister (hide) existing thumbnail if any
            if (cmdParams.SourceHwnd == IntPtr.Zero) return;

            thumbnailHandle = Thumbnail.ShowThumbnail(cmdParams.SourceHwnd, cmdParams.TargetHwnd, (Rect)cmdParams.TargetRect, out var thumbCentered);

            LogShowThumbnail(cmdParams.SourceHwnd, cmdParams.TargetHwnd, thumbCentered, thumbnailHandle);
        }

        /// <summary>
        /// Unregister (hide) the existing thumbnail identified by <see cref="MainViewModel.thumbnailHandle"/>
        /// </summary>
        /// <remarks>
        /// When there is no thumbnail (<see cref="MainViewModel.thumbnailHandle"/> is <see cref="IntPtr.Zero"/>),
        /// no exception is thrown and the method "silently" returns
        /// </remarks>
        private void HideThumbnail()
        {
            if (thumbnailHandle == IntPtr.Zero) return;

            Thumbnail.HideThumbnail(thumbnailHandle);
            LogHideThumbnail(thumbnailHandle);
            thumbnailHandle = IntPtr.Zero;
        }

        /// <summary>
        /// Launches the pinned application
        /// Parameter <paramref name="param"/> must be <see cref="PinnedAppInfo"/> object
        /// </summary>
        /// <param name="param"><see cref="PinnedAppInfo"/> object with reference to pinned application</param>
        /// <exception cref="ArgumentException">When the <paramref name="param"/> is not <see cref="PinnedAppInfo"/> object or is null, <see cref="ArgumentException"/> is thrown</exception>

        private void LaunchPinnedApp(object? param)
        {
            if (param is not PinnedAppInfo pinnedAppInfo)
            {
                LogWrongCommandParameter(nameof(PinnedAppInfo));
                throw new ArgumentException($"Command parameter must be {nameof(PinnedAppInfo)}", nameof(param));
            }

            pinnedAppInfo.LaunchPinnedApp(e =>
            {
                LogCantStartApp(pinnedAppInfo.PinnedAppType == PinnedAppTypeEnum.Package ? pinnedAppInfo.AppId ?? "[Null] appID" : pinnedAppInfo.LinkFile ?? "[Null] link file", e);
            });
        }

        /// <summary>
        /// Launches the installed application
        /// Parameter <paramref name="param"/> must be <see cref="InstalledApplication"/> object
        /// </summary>
        /// <param name="param"><see cref="InstalledApplication"/> object with reference to installed application</param>
        /// <exception cref="ArgumentException">When the <paramref name="param"/> is not <see cref="InstalledApplication"/> object or is null, <see cref="ArgumentException"/> is thrown</exception>

        private void LaunchInstalledApp(object? param)
        {
            if (param is not InstalledApplication installedApplication)
            {
                LogWrongCommandParameter(nameof(InstalledApplication));
                throw new ArgumentException($"Command parameter must be {nameof(InstalledApplication)}", nameof(param));
            }

            installedApplication.LaunchInstalledApp(e =>
            {
                LogCantStartApp(installedApplication.ShellProperties.IsStoreApp ? installedApplication.AppUserModelId ?? "[Null] appID" : installedApplication.Executable ?? "[Null] file", e);
            });
        }

        /// <summary>
        /// Builds the context menu for application window button
        /// Parameter <paramref name="param"/> must be <see cref="BuildContextMenuCommandParams"/> object
        /// </summary>
        /// <param name="param"><see cref="BuildContextMenuCommandParams"/> object with reference to <see cref="AppButton"/> and <see cref="ButtonInfo"/></param>
        /// <exception cref="ArgumentException">When the <paramref name="param"/> is not <see cref="BuildContextMenuCommandParams"/> object or is null, <see cref="ArgumentException"/> is thrown</exception>
        private void BuildContextMenu(object? param)
        {
            if (param is not BuildContextMenuCommandParams cmdParams)
            {
                LogWrongCommandParameter(nameof(BuildContextMenuCommandParams));
                throw new ArgumentException($"Command parameter must be {nameof(BuildContextMenuCommandParams)}",
                    nameof(param));
            }

            MenuItem menuItem;
            var menu = new ContextMenu();

            var buttonInfo = cmdParams.ButtonInfo;
            WndInfo? wndInfo = null;
            if (buttonInfo is WndInfo wi)
            {
                wndInfo = wi;
            }

            var isWindow = wndInfo != null;

            var appId = buttonInfo.AppId;

            if (appId != null)
            {
                //appId can be an executable full path, ensure that known folders are transformed to their GUIDs
                appId = Shell.ReplaceKnownFolderWithGuid(appId);

                //JumpList into the context menu
                var jumplistItems = jumpListService.GetJumpListItems(appId!, installedApplications);
                if (jumplistItems.Length > 0)
                {
                    string? lastCategory = null;

                    foreach (var linkInfo in jumplistItems.Where(l => l.HasTarget)) //skip separators for UI simplicity
                    {
                        if (linkInfo.Category != lastCategory)
                        {
                            //category title
                            menuItem = new MenuItem
                            {
                                Header = linkInfo.Category,
                                IsEnabled = false
                            };
                            menu.Items.Add(menuItem);

                            lastCategory = linkInfo.Category;
                        }

                        //jumplist item
                        menuItem = new MenuItem
                        {
                            Header = linkInfo.Name
                        };

                        //caption anonymization
                        if (Settings.FeatureFlag<bool>(FF_AnonymizeWindows) && linkInfo.Category != "Tasks")
                        {
                            menuItem.Header = Anonymize(linkInfo.Name, linkInfo.GetHashCode());
                        }

                        if (linkInfo.Icon != null)
                        {
                            menuItem.Icon = new Image
                            {
                                Source = Settings.InvertWhiteIcons
                                    ? Resource.InvertBitmapIfWhiteOnly(linkInfo.Icon)
                                    : linkInfo.Icon
                            };
                        }
                        else
                        {
                            //use app icon
                            menuItem.Icon = new Image
                            {
                                Source = Settings.InvertWhiteIcons
                                    ? Resource.InvertBitmapIfWhiteOnly(buttonInfo.BitmapSource)
                                    : buttonInfo.BitmapSource
                            };
                        }

                        menuItem.Click += (_, _) =>
                        {
                            try
                            {
                                if (!linkInfo.IsStoreApp)
                                {
                                    Process.Start(new ProcessStartInfo(linkInfo.TargetPath!)
                                    {
                                        Arguments = linkInfo.Arguments,
                                        WorkingDirectory = linkInfo.WorkingDirectory,
                                        UseShellExecute = true
                                    });
                                }
                                else
                                {
                                    Package.ActivateApplication(linkInfo.TargetPath, linkInfo.Arguments, out _);
                                }
                            }
                            catch (Exception ex)
                            {
                                LogCantStartApp(linkInfo.ToString(), ex);
                            }

                        };
                        menu.Items.Add(menuItem);
                    }

                    menu.Items.Add(new Separator());
                }
            }


            BuildContextMenuItemLaunchNewInstance(isWindow, buttonInfo, menu);

            if (isWindow)
            {
                //close window menu item
                menuItem = new MenuItem
                {
                    Header = "Close window",
                    Icon = new PackIconBootstrapIcons
                    {
                        Kind = PackIconBootstrapIconsKind.X,
                        Foreground = Brushes.Red
                    }
                };
                menuItem.Click += (_, _) => { WndAndApp.CloseWindow(wndInfo!.Hwnd); };
                menu.Items.Add(menuItem);
            }

            menu.Items.Add(new Separator());

            //close menu (cancel) menu item
            menuItem = new MenuItem
            {
                Header = "Cancel",
                Icon = new PackIconBootstrapIcons
                {
                    Kind = PackIconBootstrapIconsKind.Eject,
                }
            };
            menuItem.Click += (_, _) =>
            {
                //do nothing, just close the context menu
            };
            menu.Items.Add(menuItem);

            cmdParams.Button.ContextMenu = menu;
        }

        /// <summary>
        /// Builds the context menu item for launching a new instance of application
        /// </summary>
        /// <param name="isWindow">Flag whether the context menu is for window button</param>
        /// <param name="buttonInfo">Information about application window or pinned app </param>
        /// <param name="menu">Context menu</param>
        private void BuildContextMenuItemLaunchNewInstance(bool isWindow, ButtonInfo buttonInfo, ContextMenu menu)
        {
            if (isWindow)
            {
                //Start new instance menu item (window)
                if (!File.Exists(buttonInfo.Executable)) return;

                var appName =
                    installedApplications.GetInstalledApplicationFromAppId(buttonInfo.AppId ?? string.Empty)?.Name ??
                    installedApplications.GetInstalledApplicationFromExecutable(buttonInfo.Executable)?.Name ??
                    FileVersionInfo.GetVersionInfo(buttonInfo.Executable).FileDescription ??
                    Path.GetFileName(buttonInfo.Executable);

                var menuItem = new MenuItem
                {
                    Header = appName,
                    Icon = new Image
                    {
                        Source = Settings.InvertWhiteIcons
                            ? Resource.InvertBitmapIfWhiteOnly(buttonInfo.BitmapSource)
                            : buttonInfo.BitmapSource
                    }
                };
                menuItem.Click += (_, _) =>
                {
                    if (buttonInfo.Executable.ToLowerInvariant().EndsWith("\\explorer.exe"))
                    {
                        //explorer and the "special" folders like control panel
                        try
                        {
                            Process.Start(new ProcessStartInfo("explorer")
                            {
                                Arguments = buttonInfo.AppId != null
                                    ? $"shell:appsFolder\\{buttonInfo.AppId}"
                                    : null,
                            });
                        }
                        catch (Exception ex)
                        {
                            LogCantStartApp(appName, ex);
                        }
                    }
                    else
                    {
                        var started = false;
                        if (!buttonInfo.Executable.ToLowerInvariant().EndsWith("\\applicationframehost.exe"))
                        {
                            try
                            {
                                Process.Start(buttonInfo.Executable);
                                started = true;
                            }
                            catch (Exception ex)
                            {
                                LogCantStartApp(appName, ex);
                            }
                        }

                        if (started || buttonInfo.AppId == null) return;

                        //maybe store/UWP app
                        try
                        {
                            Package.ActivateApplication(buttonInfo.AppId, null, out _);
                        }
                        catch (Exception ex)
                        {
                            LogCantStartApp(appName, ex);
                        }
                    }
                };
                menu.Items.Add(menuItem);
            }
            else
            {
                //Start new instance menu item (pinned app)
                if (buttonInfo is not PinnedAppInfo pinnedAppInfo) return;
                var menuItem = new MenuItem
                {
                    Header = pinnedAppInfo.Title,
                    Icon = new Image
                    {
                        Source = Settings.InvertWhiteIcons
                            ? Resource.InvertBitmapIfWhiteOnly(pinnedAppInfo.BitmapSource)
                            : pinnedAppInfo.BitmapSource
                    }
                };
                menuItem.Click += (_, _) => { LaunchPinnedApp(pinnedAppInfo); };
                menu.Items.Add(menuItem);
            }
        }


        /// <summary>
        /// Toggles Run On Windows startup option.
        /// When it's being set, the AppSwitcherBar link is created in Windows startup folder
        /// When it's being re-set, the AppSwitcherBar link is removed from Windows startup folder
        /// </summary>
        private void ToggleRunOnWinStartup()
        {
            if (startupService.HasAppStartupLink())
            {
                startupService.RemoveAppStartupLink();
            }
            else
            {
                startupService.CreateAppStartupLink("AppSwitcherBar application");
            }

            RunOnWinStartupSet = startupService.HasAppStartupLink();
        }

        private void InitSearch()
        {
            if (!Settings.AllowSearch) return;

            //clean the search box
            SearchText = string.Empty;

            IsInSearch = true;

        }

        public ObservableCollection<SearchResultItem> SearchResults { get; private set; } = new();

        private void DoSearch(string? text)
        {
            if(!Settings.AllowSearch) return;

            if (string.IsNullOrEmpty(text))
            {
                SearchResults.Clear();
                HasSearchResults = false;
                return;
            }

            var lastDefaultRef = GetSearchResultDefault()?.ResultReference;
            SearchResults.Clear();
            var isWindowsOnlySearch = false;
            var isAppsOnlySearch = false;

            if (text.ToLowerInvariant().StartsWith("w:"))
            {
                text = (text + " ")[2..];
                isWindowsOnlySearch = true;
            }
            else if (text.ToLowerInvariant().StartsWith("a:"))
            {
                text = (text + " ")[2..];
                isAppsOnlySearch = true;
            }

            text = text.Trim();

            var categoryLimit = Settings.SearchListCategoryLimit;
            var needsSeparator = false;

            if (!isAppsOnlySearch)
            {
                var windows = ButtonManager
                    .Where(b => b is WndInfo && b.Title.Contains(text, StringComparison.InvariantCultureIgnoreCase))
                    .Cast<WndInfo>().ToArray();
                if (windows.Length > 0)
                {
                    SearchResults.Add(new SearchResultItemHeader("Windows"));
                    foreach (var window in windows.Take(categoryLimit))
                    {
                        var item = new SearchResultItemWindow(window, w => ToggleApplicationWindow(w.Hwnd, true));
                        if (window == lastDefaultRef)
                        {
                            item.IsDefault = true;
                        }

                        SearchResults.Add(item);
                    }

                    if (windows.Length > categoryLimit) SearchResults.Add(new SearchResultItemMoreItems());

                    needsSeparator = true;
                }
            }

            if (!isWindowsOnlySearch)
            {
                var pins = pinnedApplications
                    .Where(b => b.Title.Contains(text, StringComparison.InvariantCultureIgnoreCase))
                    .ToArray();
                if (pins.Length > 0)
                {
                    if (needsSeparator)
                    {
                        SearchResults.Add(new SearchResultItemSeparator());
                    }

                    SearchResults.Add(new SearchResultItemHeader("Pinned applications"));
                    foreach (var pin in pins.Take(categoryLimit))
                    {
                        var item = new SearchResultItemPinnedApp(pin, LaunchPinnedApp);
                        if (pin == lastDefaultRef)
                        {
                            item.IsDefault = true;
                        }

                        SearchResults.Add(item);
                    }

                    if (pins.Length > categoryLimit) SearchResults.Add(new SearchResultItemMoreItems());

                    needsSeparator = true;
                }


                var installs = installedApplications.SearchByName(text).ToArray();
                if (installs.Length > 0)
                {
                    if (needsSeparator)
                    {
                        SearchResults.Add(new SearchResultItemSeparator());
                    }

                    SearchResults.Add(new SearchResultItemHeader("Applications"));
                    foreach (var install in installs.Take(categoryLimit))
                    {
                        var item = new SearchResultItemInstalledApp(install, LaunchInstalledApp);
                        if (install == lastDefaultRef)
                        {
                            item.IsDefault = true;
                        }

                        SearchResults.Add(item);
                    }

                    if (installs.Length > categoryLimit) SearchResults.Add(new SearchResultItemMoreItems());
                }
            }

            HasSearchResults = SearchResults.Count > 0;
            if (GetSearchResultDefault() == null && HasSearchResults)
            {
                GetSearchResultsWithRef()[0].IsDefault = true;
            }
        }

        private void EndSearch()
        {
            //clean the search box
            SearchText = string.Empty;
            IsInSearch = false;
        }

        /// <summary>
        /// Process the special key for search
        /// Parameter <paramref name="param"/> must be <see cref="Key"/> value
        /// </summary>
        /// <param name="param"><see cref="Key"/> pressed </param>
        /// <exception cref="ArgumentException">When the <paramref name="param"/> is not <see cref="Key"/> object or is null, <see cref="ArgumentException"/> is thrown</exception>
        private void SearchSpecialKey(object? param)
        {
            if (!Settings.AllowSearch) return;

            if (param is not Key key)
            {
                LogWrongCommandParameter(nameof(Key));
                throw new ArgumentException($"Command parameter must be {nameof(Key)}",
                    nameof(param));
            }

            if (key == Key.Escape)
            {
                if (!string.IsNullOrEmpty(SearchText))
                {
                    SearchText = string.Empty;
                }
                else
                {
                    EndSearch();
                }
            }

            if (key == Key.Enter)
            {
                GetSearchResultDefault()?.Launch();
            }

            if (key == Key.Up)
            {
                var results = GetSearchResultsWithRef();
                var current = GetSearchResultDefault();
                if (current != null)
                {
                    var idx = results.IndexOf(current);
                    idx--;
                    if (idx >= 0)
                    {
                        current.IsDefault = false;
                        results[idx].IsDefault = true;
                    }
                }
            }

            if (key == Key.Down)
            {
                var results = GetSearchResultsWithRef();
                var current = GetSearchResultDefault();
                if (current != null)
                {
                    var idx = results.IndexOf(current);
                    idx++;
                    if (idx < results.Count)
                    {
                        current.IsDefault = false;
                        results[idx].IsDefault = true;
                    }
                }
            }

            if (key == Key.PageUp)
            {
                var results = GetSearchResultsWithRef();
                var current = GetSearchResultDefault();

                if (current != null)
                {
                    var idx = results.IndexOf(current);
                    var currentType = current.GetType();

                    do
                    {
                        idx--;
                        if (idx < 0 || results[idx].GetType() == currentType) continue;

                        current.IsDefault = false;
                        results[idx].IsDefault = true;
                        break;
                    } while (idx >= 0);
                }
            }

            if (key == Key.PageDown)
            {
                var results = GetSearchResultsWithRef();
                var current = GetSearchResultDefault();

                if (current != null)
                {
                    var idx = results.IndexOf(current);
                    var currentType = current.GetType();

                    do
                    {
                        idx++;
                        if (idx >= results.Count || results[idx].GetType() == currentType) continue;

                        current.IsDefault = false;
                        results[idx].IsDefault = true;
                        break;
                    } while (idx < results.Count);
                }
            }
        }

        private List<SearchResultItemWithRef> GetSearchResultsWithRef()
        {
            var results = SearchResults.Where(r => r is SearchResultItemWithRef).Cast<SearchResultItemWithRef>().ToList();
            return results;
        }
        private SearchResultItemWithRef? GetSearchResultDefault()
        {
            var result = GetSearchResultsWithRef().FirstOrDefault(r => r.IsDefault);
            return result;
        }

        /// <summary>
        /// Occurs when a property value changes
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// Raise <see cref="PropertyChanged"/> event for given <paramref name="propertyName"/>
        /// </summary>
        /// <param name="propertyName">Name of the property changed</param>
        // ReSharper disable once UnusedMember.Global
        public void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }



}
