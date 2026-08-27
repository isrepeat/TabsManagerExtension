using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.ExtensionManager;
using Task = System.Threading.Tasks.Task;

// Add AUTO_ENABLE_CUSTOM_TABS to DefineConstants to restore automatic left layout
// and tab replacement. Keep it undefined while diagnosing manual activation.

[assembly: Helpers.Attributes.CodeAnalyzerEnableLogs]

#if NETFRAMEWORK
namespace System.Runtime.CompilerServices {
    internal static class IsExternalInit { } // need for "init" keyword
}
#endif

namespace TabsManagerExtension {
    /// <summary>
    /// ProvideAutoLoad просит Visual Studio автоматически загрузить расширение в фоновом режиме.
    /// NoSolution действует, когда решение не открыто, а SolutionExists — когда решение открыто.
    /// Поэтому расширение загружается в обоих возможных состояниях Visual Studio без старого хака с Tool Window.
    /// </summary>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
#if ENABLE_EARLY_PACKAGE_LOAD_HACK
    [ProvideToolWindow(typeof(ToolWindows.EarlyPackageLoadHackToolWindow))]
#endif
    [ProvideToolWindow(typeof(ToolWindows.TabsManagerToolWindow))]
    [ProvideToolWindow(typeof(ToolWindows.TabsManagerSettingsToolWindow))]
    [Guid(TabsManagerExtensionPackage.PackageGuidString)]
    public sealed class TabsManagerExtensionPackage : AsyncPackage {
        public const string PackageGuidString = "7a0ce045-e2ba-4f14-8b80-55cfd666e3d8";
        // Идентификатор VSIX из source.extension.vsixmanifest. Extension Manager использует
        // его для поиска именно установленного экземпляра всего расширения.
        private const string ExtensionId = "TabsManagerExtension.93937105-d7bb-4eee-9a06-2eb9d8353aab";
        private static TabsManagerExtensionPackage? _instance;

        // Версия берётся из Header установленного VSIX, а не дублируется в коде UI.
        // TryGetInstalledExtension также безопасен для временно недоступного пакета.
        internal static string GetInstalledExtensionVersion() {
            ThreadHelper.ThrowIfNotOnUIThread();
            try {
                var extensionManager = Package.GetGlobalService(typeof(SVsExtensionManager)) as IVsExtensionManager;
                if (extensionManager != null &&
                    extensionManager.TryGetInstalledExtension(ExtensionId, out var extension) &&
                    extension?.Header?.Version != null) {

                    return extension.Header.Version.ToString();
                }
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogWarning($"Failed to get installed Tabs Manager extension version: {ex.Message}");
            }

            return "unknown";
        }

        // Синхронная точка входа используется командами VS и передаёт работу в JTF без блокировки UI.
        internal static void ShowOptions() {
            ThreadHelper.ThrowIfNotOnUIThread();
            ThreadHelper.JoinableTaskFactory.RunAsync(ShowCustomSettingsAsync).FileAndForget("TabsManagerExtension/ShowSettings");
        }

        internal static async Task ShowCustomSettingsAsync() {
            // Страница реализована как MDI tool window: это даёт полный контроль над XAML и поведение вкладки документа.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            Helpers.Diagnostic.Logger.LogDebug("[Settings] Opening Tabs Manager settings tab.");
            var package = _instance ?? throw new InvalidOperationException("Tabs Manager package is not initialized.");
            var window = await package.FindToolWindowAsync(
                typeof(ToolWindows.TabsManagerSettingsToolWindow),
                0,
                true,
                package.DisposalToken
            );

            if (window?.Frame is not IVsWindowFrame frame) {
                throw new InvalidOperationException("Unable to create Tabs Manager settings window.");
            }

            frame.SetProperty(
                (int)__VSFPROPID.VSFPROPID_FrameMode,
                (int)VSFRAMEMODE.VSFM_MdiChild
            );

            // Show активирует существующий фрейм либо показывает только что созданный.
            ErrorHandler.ThrowOnFailure(frame.Show());
            Helpers.Diagnostic.Logger.LogDebug("[Settings] Tabs Manager settings tab opened.");
        }

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress) {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            _instance = this;

            string logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TabsManagerExtension"
            );
            string logFilePath = Path.Combine(logDirectory, "TabsManagerExtension.log");
            Helpers.Diagnostic.Logger.EnableFileLogging(logFilePath);
            Helpers.Diagnostic.Logger.LogDebug($"[Package] File logging enabled: '{logFilePath}'.");

            //Console.Beep(1000, 500); // 1000 Гц, 500 мс
            Settings.TabsManagerSettingsService.SettingsInitialized += this.OnSettingsInitialized;
            Settings.TabsManagerSettingsService.AutoLoadCustomTabsChanged += this.OnAutoLoadCustomTabsChanged;

            Services.ExtensionServices.Initialize();

            // По умолчанию вкладки включены. Встраиваем их сразу, не ожидая ленивой
            // активации VisualStudio.Extensibility и чтения пользовательских настроек.
            this.EnableCustomTabsImmediately();

#if DEBUG
            this.ConfigureExperimentalStartup();
            this.OpenDebugSolutionIfNoneLoaded();
#endif
            this.InitializeEvents();

#if ENABLE_EARLY_PACKAGE_LOAD_HACK
            ToolWindows.EarlyPackageLoadHackToolWindow.Initialize(this);
#endif
            if (Settings.TabsManagerSettingsService.IsSettingsInitialized) {
                this.ApplyCustomTabsSetting();
            }

            await ToolWindows.TabsManagerToolWindowCommand.InitializeAsync(this);
            this.ShowLoadedStatusAfterShellBecomesIdle();
        }

#if DEBUG
        // Fallback используется только экспериментальной DEBUG-средой и не влияет на VSIX Release.
        private const string DebugSolutionPath = @"C:\WORK\Projects\TabsManagerExtension - Copy\TabsManagerExtension.sln";
        //private const string DebugSolutionPath = @"C:\WORK\Projects\Cpp\UtilityHelpersLib\UtilityHelpersLib.sln";

        private void ConfigureExperimentalStartup() {
            ThreadHelper.ThrowIfNotOnUIThread();

            try {
                var startupProperties = PackageServices.Dte2.Properties["Environment", "Startup"];
                var onStartup = startupProperties?.Item("OnStartUp");
                int loadLastSolution = (int)EnvDTE.vsStartUp.vsStartUpLoadLastSolution;
                if (onStartup != null && Convert.ToInt32(onStartup.Value) != loadLastSolution) {
                    onStartup.Value = loadLastSolution;
                }
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogDebug($"[Package] Failed to configure Experimental Instance startup: {ex.Message}");
            }
        }

        private void OpenDebugSolutionIfNoneLoaded() {
            // Запускаем проверку в фоне, чтобы инициализация пакета не блокировала старт IDE.
            this.JoinableTaskFactory.RunAsync(async () => {
                // Пакет может загрузиться в кратком состоянии NoSolution до обработки аргументов devenv.
                // Даём штатному запуску время открыть переданное решение и вмешиваемся только в пустую IDE.
                await Task.Delay(TimeSpan.FromSeconds(1), this.DisposalToken);
                await this.JoinableTaskFactory.SwitchToMainThreadAsync(this.DisposalToken);

                try {
                    if (PackageServices.Dte2.Solution.IsOpen) {
                        return;
                    }

                    if (!System.IO.File.Exists(DebugSolutionPath)) {
                        Helpers.Diagnostic.Logger.LogDebug($"[Package] Debug solution was not found: '{DebugSolutionPath}'.");
                        return;
                    }

                    Helpers.Diagnostic.Logger.LogDebug($"[Package] Experimental Instance started without a solution. Opening '{DebugSolutionPath}'.");
                    PackageServices.Dte2.Solution.Open(DebugSolutionPath);
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogDebug($"[Package] Failed to open the debug solution: {ex.Message}");
                }
            }).FileAndForget("TabsManagerExtension/OpenDebugSolutionIfNoneLoaded");
        }
#endif

        private void OnSettingsInitialized() {
            this.JoinableTaskFactory.RunAsync(async () => {
                await this.JoinableTaskFactory.SwitchToMainThreadAsync(this.DisposalToken);
                this.ApplyCustomTabsSetting();
            }).FileAndForget("TabsManagerExtension/RestoreCustomTabsAfterSettingsInitialization");
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                Settings.TabsManagerSettingsService.SettingsInitialized -= this.OnSettingsInitialized;
                Settings.TabsManagerSettingsService.AutoLoadCustomTabsChanged -= this.OnAutoLoadCustomTabsChanged;
                Settings.TabsManagerSettingsService.Shutdown();
            }

            base.Dispose(disposing);
        }

        private void EnableCustomTabsImmediately() {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!Settings.TabsManagerSettingsService.AutoLoadCustomTabs) {
                return;
            }

            if (!VsixVisualTreeHelper.Instance.IsCustomTabsEnabled) {
                Helpers.Diagnostic.Logger.LogDebug("[Package] Enabling custom tabs early according to local settings.json.");
                VsixVisualTreeHelper.Instance.ToggleCustomTabs(true, savePreference: false);
            }
        }

        private void OnAutoLoadCustomTabsChanged(bool enabled) {
            this.JoinableTaskFactory.RunAsync(async () => {
                await this.JoinableTaskFactory.SwitchToMainThreadAsync(this.DisposalToken);
                VsixVisualTreeHelper.Instance.ToggleCustomTabs(enabled, savePreference: false);
            }).FileAndForget("TabsManagerExtension/ApplyTabsSwitch");
        }

        private void ApplyCustomTabsSetting() {
            ThreadHelper.ThrowIfNotOnUIThread();

            bool shouldEnable = Settings.TabsManagerSettingsService.AutoLoadCustomTabs;
            if (shouldEnable) {
                // Даже если логическое состояние уже включено, host мог отсутствовать
                // на Start Window и появиться только после открытия решения.
                VsixVisualTreeHelper.Instance.ToggleCustomTabs(true, savePreference: false);
                return;
            }

            if (VsixVisualTreeHelper.Instance.IsCustomTabsEnabled) {
                Helpers.Diagnostic.Logger.LogDebug("[Package] Applying AutoLoadCustomTabs=False.");
                VsixVisualTreeHelper.Instance.ToggleCustomTabs(false, savePreference: false);
            }
        }


        private void InitializeEvents() {
            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.SolutionLoaded.Add(this.OnSolutionLoaded);
            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.SolutionLoaded.InvokeForLastHandlerIfTriggered();

            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.SolutionClosed.Add(this.OnSolutionClosed);
            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.SolutionClosed.InvokeForLastHandlerIfTriggered();
        }


        private void OnSolutionLoaded(string solutionName) {
            Helpers.Diagnostic.Logger.LogDebug($"[Package] OnSolutionLoaded(): solutionName = {solutionName}");
            PackageServices.Invalidate();
            this.ShowLoadedStatusAfterShellBecomesIdle();

            if (VsixVisualTreeHelper.Instance.IsCustomTabsEnabled) {
                // На Start Window области документов ещё нет. После загрузки решения
                // повторяем внедрение в уже созданный PART_TabListHost.
                VsixVisualTreeHelper.Instance.ToggleCustomTabs(true, savePreference: false);
            }

#if AUTO_ENABLE_CUSTOM_TABS
            if (VsixVisualTreeHelper.Instance.IsCustomTabsEnabled) {
                return;
            }

            VsixThreadHelper.RunOnVsThread(() => {
                VsixVisualTreeHelper.Instance.ToggleCustomTabs(true);
            });
#endif
        }

        private void ShowLoadedStatusAfterShellBecomesIdle() {
            this.JoinableTaskFactory.RunAsync(async () => {
                // Solution loading writes its own status messages after package initialization.
                // Wait until those messages finish so the activation confirmation remains visible.
                await Task.Delay(TimeSpan.FromSeconds(2), this.DisposalToken);
                await this.JoinableTaskFactory.SwitchToMainThreadAsync(this.DisposalToken);

                var statusBar = await this.GetServiceAsync(typeof(SVsStatusbar)) as IVsStatusbar;
                statusBar?.SetText("Tabs Manager loaded — use Tools > Toggle Tabs Manager");
            }).FileAndForget("TabsManagerExtension/ShowLoadedStatus");
        }


        private void OnSolutionClosed(string solutionName) {
            Helpers.Diagnostic.Logger.LogDebug($"[Package] OnSolutionClosed(): solutionName = {solutionName}");
            PackageServices.Invalidate();

#if AUTO_ENABLE_CUSTOM_TABS
            if (!VsixVisualTreeHelper.Instance.IsCustomTabsEnabled) {
                return;
            }

            VsixThreadHelper.RunOnVsThread(() => {
                VsixVisualTreeHelper.Instance.ToggleCustomTabs(false);
            });
#endif
        }
    }
}


namespace TabsManagerExtension.Behaviours {
    public class Dummy {
    }
}
