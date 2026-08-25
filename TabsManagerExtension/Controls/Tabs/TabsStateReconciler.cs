using System;
using System.Linq;
using System.Windows.Threading;
using System.Collections.Generic;

using Microsoft.VisualStudio.Shell;

using TabsManagerExtension.State.Document;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Периодически сверяет модели вкладок с фактическим состоянием Visual Studio.</summary>
    internal sealed class TabsStateReconciler : IDisposable {
        private readonly EnvDTE80.DTE2 _dte;
        private readonly DispatcherTimer _timer;
        private readonly TabCollectionManager _tabCollectionManager;
        private readonly Action _updateWindowTabsInfo;

        public TabsStateReconciler(
            EnvDTE80.DTE2 dte,
            Dispatcher dispatcher,
            TabCollectionManager tabCollectionManager,
            Action updateWindowTabsInfo
            ) {
            _dte = dte;
            _tabCollectionManager = tabCollectionManager;
            _updateWindowTabsInfo = updateWindowTabsInfo;
            _timer = new DispatcherTimer(DispatcherPriority.Normal, dispatcher) {
                Interval = TimeSpan.FromSeconds(2)
            };
            _timer.Tick += this.OnTimerTick;
        }

        public void Start() {
            _timer.Start();
        }

        public void Stop() {
            _timer.Stop();
        }

        public void Reconcile() {
            ThreadHelper.ThrowIfNotOnUIThread();

            this.UpdateDocumentSaveStates();
            this.MoveDocumentsOutOfPreviewGroup();
            this.RemoveClosedToolWindows();
            _updateWindowTabsInfo();
        }

        public void Dispose() {
            _timer.Stop();
            _timer.Tick -= this.OnTimerTick;
        }

        private void OnTimerTick(object sender, EventArgs e) {
            this.Reconcile();
        }

        private void UpdateDocumentSaveStates() {
            var openDocuments = _dte.Documents
                .Cast<EnvDTE.Document>()
                .ToDictionary(document => document.FullName, StringComparer.OrdinalIgnoreCase);

            foreach (var group in _tabCollectionManager.Groups.ToList()) {
                foreach (var tabItem in group.Items.ToList()) {
                    if (!openDocuments.TryGetValue(tabItem.FullName, out var document)) {
                        continue;
                    }

                    if (document.Saved) {
                        tabItem.Caption = tabItem.Caption.TrimEnd('*');
                    }
                    else if (!tabItem.Caption.EndsWith("*")) {
                        tabItem.Caption += "*";
                    }
                }
            }
        }

        private void MoveDocumentsOutOfPreviewGroup() {
            var previewGroup = _tabCollectionManager.Groups.FirstOrDefault(group => group is TabItemsPreviewGroup);
            if (previewGroup == null) {
                return;
            }

            foreach (var document in previewGroup.Items.OfType<TabItemDocument>().ToList()) {
                if (!document.ShellDocument.IsDocumentInPreviewTab()) {
                    _tabCollectionManager.MovePreviewDocumentToDefaultGroup(document);
                }
            }
        }

        private void RemoveClosedToolWindows() {
            var openWindowIds = new HashSet<string>();
            try {
                openWindowIds = _dte.Windows
                    .Cast<EnvDTE.Window>()
                    .Select(VsShell.Document.ShellWindow.GetWindowId)
                    .Where(windowId => !string.IsNullOrEmpty(windowId))
                    .ToHashSet();
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to enumerate windows: {ex.Message}");
            }

            foreach (var group in _tabCollectionManager.Groups.ToList()) {
                var windowsToRemove = group.Items
                    .OfType<TabItemWindow>()
                    .Where(window => !openWindowIds.Contains(window.WindowId) && !IsWindowVisible(window))
                    .ToList();

                foreach (var window in windowsToRemove) {
                    _tabCollectionManager.Remove(window);
                }
            }
        }

        private static bool IsWindowVisible(TabItemWindow tabItemWindow) {
            ThreadHelper.ThrowIfNotOnUIThread();

            try {
                return tabItemWindow.ShellWindow.Window.Visible;
            }
            catch (Exception) {
                return false;
            }
        }
    }
}
