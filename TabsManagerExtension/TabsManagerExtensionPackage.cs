using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.ComponentModel.Design;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.VCProjectEngine;
using Microsoft.VisualStudio.VCCodeModel;
using Microsoft.VisualStudio.TextManager.Interop;
using Task = System.Threading.Tasks.Task;

// Add AUTO_ENABLE_CUSTOM_TABS to DefineConstants to restore automatic left layout
// and tab replacement. Keep it undefined while diagnosing manual activation.

[assembly: Helpers.Attributes.CodeAnalyzerEnableLogs]

#if NET_FRAMEWORK_472
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
    [Guid(TabsManagerExtensionPackage.PackageGuidString)]
    public sealed class TabsManagerExtensionPackage : AsyncPackage {
        public const string PackageGuidString = "7a0ce045-e2ba-4f14-8b80-55cfd666e3d8";
        private const string OptionKey = "TabsManagerExtension";

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress) {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            
            // TODO: adapt the CppFeatures nuget to net472 (WPF?)
            //var initFlags = CppFeatures.Cx.InitFlags.DefaultFlags | CppFeatures.Cx.InitFlags.CreateInPackageFolder;
            //CppFeatures.Cx.Logger.Init(AppConstants.LogFilename, initFlags);

            //Console.Beep(1000, 500); // 1000 Гц, 500 мс
            Services.ExtensionServices.Initialize();

            this.InitializeEvents();

#if ENABLE_EARLY_PACKAGE_LOAD_HACK
            ToolWindows.EarlyPackageLoadHackToolWindow.Initialize(this);
#endif
            await ToolWindows.TabsManagerToolWindowCommand.InitializeAsync(this);
            this.ShowLoadedStatusAfterShellBecomesIdle();
            this.RestoreCustomTabsAfterShellBecomesIdle();
        }

        private void RestoreCustomTabsAfterShellBecomesIdle() {
            if (!Configuration.TabsManagerConfigurationService.AutoLoadCustomTabs) {
                return;
            }

            this.JoinableTaskFactory.RunAsync(async () => {
                await Task.Delay(TimeSpan.FromSeconds(2), this.DisposalToken);
                await this.JoinableTaskFactory.SwitchToMainThreadAsync(this.DisposalToken);

                if (!VsixVisualTreeHelper.Instance.IsCustomTabsEnabled) {
                    VsixVisualTreeHelper.Instance.ToggleCustomTabs(true);
                }
            }).FileAndForget("TabsManagerExtension/RestoreCustomTabs");
        }


        private void InitializeEvents() {
            VsShell.Services.VsIDEStateFlagsTrackerService.Instance.SolutionLoaded.Add(this.OnSolutionLoaded);
            VsShell.Services.VsIDEStateFlagsTrackerService.Instance.SolutionLoaded.InvokeForLastHandlerIfTriggered();

            VsShell.Services.VsIDEStateFlagsTrackerService.Instance.SolutionClosed.Add(this.OnSolutionClosed);
            VsShell.Services.VsIDEStateFlagsTrackerService.Instance.SolutionClosed.InvokeForLastHandlerIfTriggered();
        }


        private void OnSolutionLoaded(string solutionName) {
            Helpers.Diagnostic.Logger.LogDebug($"[Package] OnSolutionLoaded(): solutionName = {solutionName}");
            PackageServices.Invalidate();
            this.ShowLoadedStatusAfterShellBecomesIdle();

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
