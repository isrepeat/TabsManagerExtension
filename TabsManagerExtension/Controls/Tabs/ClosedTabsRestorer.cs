using System;
using System.IO;
using System.Linq;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using TabsManagerExtension.State.Document;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Открывает закрытые вкладки и возвращает их в сохранённые группы.</summary>
    internal sealed class ClosedTabsRestorer {
        private readonly EnvDTE80.DTE2 _dte;
        private readonly TabCollectionManager _tabCollectionManager;
        private readonly Action _updateWindowTabsInfo;
        private readonly Action _restoreInputTarget;
        private readonly Action _focusInputTarget;

        public bool IsRestoring { get; private set; }

        public ClosedTabsRestorer(
            EnvDTE80.DTE2 dte,
            TabCollectionManager tabCollectionManager,
            Action updateWindowTabsInfo,
            Action restoreInputTarget,
            Action focusInputTarget
            ) {
            _dte = dte;
            _tabCollectionManager = tabCollectionManager;
            _updateWindowTabsInfo = updateWindowTabsInfo;
            _restoreInputTarget = restoreInputTarget;
            _focusInputTarget = focusInputTarget;
        }

        public void Restore(ClosedTabsOperation operation) {
            ThreadHelper.ThrowIfNotOnUIThread();

            this.IsRestoring = true;
            try {
                foreach (var entry in operation.Entries) {
                    try {
                        TabItemBase? restoredTabItem = entry.Kind == ClosedTabKind.Document
                            ? this.RestoreDocument(entry)
                            : this.RestoreToolWindow(entry);

                        if (restoredTabItem != null) {
                            this.MoveToOriginalGroup(restoredTabItem, entry);
                        }
                    }
                    catch (Exception ex) {
                        Helpers.Diagnostic.Logger.LogError($"Failed to restore closed tab '{entry.FullName}': {ex}");
                    }
                }
            }
            finally {
                this.IsRestoring = false;
            }

            _focusInputTarget();
            _restoreInputTarget();
        }

        private TabItemDocument? RestoreDocument(ClosedTabEntry entry) {
            var existingTabItem = _tabCollectionManager.Find(entry.FullName);
            if (existingTabItem != null) {
                return existingTabItem;
            }
            if (!File.Exists(entry.FullName)) {
                Helpers.Diagnostic.Logger.LogWarning($"Cannot restore deleted document '{entry.FullName}'");
                return null;
            }

            var window = _dte.ItemOperations.OpenFile(entry.FullName);
            var restoredTabItem = window?.Document == null ? null : _tabCollectionManager.Find(window.Document);
            return restoredTabItem ?? _tabCollectionManager.Find(entry.FullName);
        }

        private TabItemWindow? RestoreToolWindow(ClosedTabEntry entry) {
            if (!Guid.TryParse(entry.WindowId, out var persistenceGuid)) {
                return null;
            }

            var uiShell = Package.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
            if (uiShell == null) {
                return null;
            }

            int result = uiShell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fForceCreate, ref persistenceGuid, out var frame);
            if (ErrorHandler.Failed(result) || frame == null) {
                return null;
            }

            frame.Show();
            _updateWindowTabsInfo();
            return _tabCollectionManager.Groups
                .SelectMany(group => group.Items)
                .OfType<TabItemWindow>()
                .FirstOrDefault(item => string.Equals(item.WindowId, entry.WindowId, StringComparison.OrdinalIgnoreCase));
        }

        private void MoveToOriginalGroup(TabItemBase tabItem, ClosedTabEntry entry) {
            var current = _tabCollectionManager.FindWithGroup(tabItem);
            if (current != null) {
                _tabCollectionManager.RemoveFromGroup(current.Value.Item, current.Value.Group);
            }

            _tabCollectionManager.AddToGroup(tabItem, CreateGroup(entry));
            if (tabItem is TabItemDocument tabItemDocument && entry.GroupKind == ClosedTabGroupKind.Pinned) {
                tabItemDocument.ShellDocument.OpenDocumentAsPinned();
            }
        }

        private static TabItemsGroupBase CreateGroup(ClosedTabEntry entry) {
            return entry.GroupKind switch {
                ClosedTabGroupKind.Preview => new TabItemsPreviewGroup(),
                ClosedTabGroupKind.Pinned => new TabItemsPinnedGroup(entry.GroupName),
                _ => new TabItemsDefaultGroup(entry.GroupName)
            };
        }
    }
}
