using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Diagnostics;
using System.Collections.Generic;

using Microsoft.VisualStudio.Shell;

using TabsManagerExtension.State.Document;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Выполняет пользовательские команды над одной или несколькими вкладками.</summary>
    internal sealed class TabCommandController {
        private readonly VirtualMenuControl _virtualMenu;
        private readonly TabCollectionManager _tabCollectionManager;
        private readonly ClosedTabsHistory _closedTabsHistory;
        private readonly Helpers.Collections.GroupsSelectionCoordinator<TabItemsGroupBase, TabItemBase> _selectionCoordinator;
        private readonly Func<TabItemBase, ClosedTabEntry> _createClosedTabEntry;

        public TabCommandController(
            VirtualMenuControl virtualMenu,
            TabCollectionManager tabCollectionManager,
            ClosedTabsHistory closedTabsHistory,
            Helpers.Collections.GroupsSelectionCoordinator<TabItemsGroupBase, TabItemBase> selectionCoordinator,
            Func<TabItemBase, ClosedTabEntry> createClosedTabEntry
            ) {
            _virtualMenu = virtualMenu;
            _tabCollectionManager = tabCollectionManager;
            _closedTabsHistory = closedTabsHistory;
            _selectionCoordinator = selectionCoordinator;
            _createClosedTabEntry = createClosedTabEntry;
        }

        public void Pin(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is not TabItemBase tabItem || tabItem.IsPinnedTab) {
                return;
            }

            var current = _tabCollectionManager.FindWithGroup(tabItem);
            if (current != null) {
                _tabCollectionManager.Move(tabItem, new TabItemsPinnedGroup(current.Value.Group.GroupName));
            }
        }

        public void Unpin(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is not TabItemBase tabItem || !tabItem.IsPinnedTab) {
                return;
            }

            var current = _tabCollectionManager.FindWithGroup(tabItem);
            if (current != null) {
                _tabCollectionManager.Move(tabItem, new TabItemsDefaultGroup(current.Value.Group.GroupName));
            }
        }

        public void Close(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is not TabItemBase tabItem) {
                return;
            }

            var selectedItems = _selectionCoordinator.SelectedItems;
            bool closeSelection = selectedItems.Count > 1 &&
                selectedItems.Any(entry => ReferenceEquals(entry.Item, tabItem));
            var itemsToClose = closeSelection
                ? selectedItems.Select(entry => entry.Item).ToList()
                : new List<TabItemBase> { tabItem };
            this.Close(itemsToClose);
        }

        public void CloseSelected() {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.Close(_selectionCoordinator.SelectedItems.Select(entry => entry.Item).ToList());
        }

        public void Close(IReadOnlyList<TabItemBase> tabItems) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Снимок нужен до DTE Close(): DocumentClosing приходит синхронно и удаляет модель.
            var closeRequests = tabItems
                .Select(tabItem => new {
                    TabItem = tabItem,
                    Entry = _createClosedTabEntry(tabItem)
                })
                .ToList();
            foreach (var request in closeRequests) {
                _closedTabsHistory.MarkClosing(request.TabItem);
            }

            var closedEntries = new List<ClosedTabEntry>();
            foreach (var request in closeRequests) {
                var tabItem = request.TabItem;
                try {
                    if (tabItem is TabItemDocument tabItemDocument) {
                        Helpers.Diagnostic.Logger.LogDebug($"close document \"{tabItemDocument.ShellDocument.Document.FullName}\"");
                        tabItemDocument.ShellDocument.Close();
                    }
                    else if (tabItem is TabItemWindow tabItemWindow) {
                        Helpers.Diagnostic.Logger.LogDebug($"close window \"{tabItemWindow.ShellWindow.Window.Caption}\"");
                        tabItemWindow.ShellWindow.Window.Close();
                        _tabCollectionManager.Remove(tabItemWindow);
                    }
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogError($"Failed to close tab '{tabItem.Caption}': {ex}");
                }
                finally {
                    if (_tabCollectionManager.FindWithGroup(tabItem) == null) {
                        closedEntries.Add(request.Entry);
                    }
                    else {
                        Helpers.Diagnostic.Logger.LogWarning($"Tab '{tabItem.Caption}' remained open after close request");
                    }

                    _closedTabsHistory.RemoveClosingMark(tabItem);
                }
            }

            if (closedEntries.Count > 0) {
                _closedTabsHistory.Push(closedEntries);
            }

            _virtualMenu.HideImmediately();
        }

        public void KeepOpen(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is TabItemDocument tabItemDocument) {
                _tabCollectionManager.MovePreviewDocumentToDefaultGroup(tabItemDocument);
                tabItemDocument.ShellDocument.OpenDocumentAsPinned();
            }
        }

        public void CopyName(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is TabItemBase tabItem) {
                this.CopyText(tabItem.Caption, "name");
            }
        }

        public void CopyPath(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is TabItemBase tabItem) {
                this.CopyText(tabItem.FullName, "path");
            }
        }

        public void OpenLocation(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is TabItemDocument tabItemDocument) {
                try {
                    string filePath = tabItemDocument.FullName;
                    if (File.Exists(filePath)) {
                        Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                    }
                    else {
                        Helpers.Diagnostic.Logger.LogWarning($"File not found: {filePath}");
                    }
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogError($"Failed to open tab location: {ex.Message}");
                }
            }

            _virtualMenu.HideImmediately();
        }

        private void CopyText(string text, string valueKind) {
            ThreadHelper.ThrowIfNotOnUIThread();
            try {
                Clipboard.SetText(text ?? string.Empty);
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to copy tab {valueKind} to clipboard: {ex}");
            }
        }
    }
}
