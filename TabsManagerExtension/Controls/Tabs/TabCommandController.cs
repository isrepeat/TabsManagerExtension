using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Diagnostics;
using System.Collections.Generic;
using Microsoft.VisualStudio.Shell;

using TMEx = TabsManagerExtension;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Выполняет пользовательские команды над одной или несколькими вкладками.</summary>
    internal sealed class TabCommandController {
        private readonly VirtualMenuControl _virtualMenu;
        private readonly TabCollectionManager _tabCollectionManager;
        private readonly ClosedTabsHistory _closedTabsHistory;
        private readonly Helpers.Collections.GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase> _selectionCoordinator;
        private readonly Func<TMEx.State.Document.TabItemBase, ClosedTabEntry> _createClosedTabEntry;

        public TabCommandController(
            VirtualMenuControl virtualMenu,
            TabCollectionManager tabCollectionManager,
            ClosedTabsHistory closedTabsHistory,
            Helpers.Collections.GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase> selectionCoordinator,
            Func<TMEx.State.Document.TabItemBase, ClosedTabEntry> createClosedTabEntry
            ) {
            _virtualMenu = virtualMenu;
            _tabCollectionManager = tabCollectionManager;
            _closedTabsHistory = closedTabsHistory;
            _selectionCoordinator = selectionCoordinator;
            _createClosedTabEntry = createClosedTabEntry;
        }

        public void Pin(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is not TMEx.State.Document.TabItemBase tabItem || tabItem.IsPinnedTab) {
                return;
            }

            var current = _tabCollectionManager.FindWithGroup(tabItem);
            if (current != null) {
                // Меняется только секция, логическое имя группы сохраняется: после Unpin
                // вкладка вернётся в default-группу того же проекта.
                _tabCollectionManager.Move(tabItem, new TMEx.State.Document.TabItemsPinnedGroup(current.Value.Group.GroupName));
            }
        }

        public void Unpin(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is not TMEx.State.Document.TabItemBase tabItem || !tabItem.IsPinnedTab) {
                return;
            }

            var current = _tabCollectionManager.FindWithGroup(tabItem);
            if (current != null) {
                _tabCollectionManager.Move(tabItem, new TMEx.State.Document.TabItemsDefaultGroup(current.Value.Group.GroupName));
            }
        }

        public void Close(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is not TMEx.State.Document.TabItemBase tabItem) {
                return;
            }

            var selectedItems = _selectionCoordinator.SelectedItems;
            // Контекстная команда над одной из выбранных вкладок применяется ко всему
            // мультивыбору. Клик по вкладке вне selection закрывает только её.
            bool closeSelection = selectedItems.Count > 1 &&
                selectedItems.Any(entry => ReferenceEquals(entry.Item, tabItem));
            var itemsToClose = closeSelection
                ? selectedItems.Select(entry => entry.Item).ToList()
                : new List<TMEx.State.Document.TabItemBase> { tabItem };
            this.Close(itemsToClose);
        }

        public void CloseSelected() {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.Close(_selectionCoordinator.SelectedItems.Select(entry => entry.Item).ToList());
        }

        public void Close(IReadOnlyList<TMEx.State.Document.TabItemBase> tabItems) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Снимок нужен до DTE Close(): DocumentClosing приходит синхронно и удаляет модель.
            var closeRequests = tabItems
                .Select(tabItem => new {
                    TabItem = tabItem,
                    Entry = _createClosedTabEntry(tabItem)
                })
                .ToList();
            foreach (var request in closeRequests) {
                // Маркер не даёт синхронному DocumentClosing добавить каждую вкладку в history
                // отдельно. После завершения команды ниже будет создана одна undo-операция.
                _closedTabsHistory.MarkClosing(request.TabItem);
            }

            var closedEntries = new List<ClosedTabEntry>();
            foreach (var request in closeRequests) {
                var tabItem = request.TabItem;
                try {
                    if (tabItem is TMEx.State.Document.TabItemDocument tabItemDocument) {
                        Helpers.Diagnostic.Logger.LogDebug($"close document \"{tabItemDocument.ShellDocument.Document.FullName}\"");
                        tabItemDocument.ShellDocument.Close();
                    }
                    else if (tabItem is TMEx.State.Document.TabItemWindow tabItemWindow) {
                        Helpers.Diagnostic.Logger.LogDebug($"close window \"{tabItemWindow.ShellWindow.Window.Caption}\"");
                        tabItemWindow.ShellWindow.Window.Close();
                        _tabCollectionManager.Remove(tabItemWindow);
                    }
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogError($"Failed to close tab '{tabItem.Caption}': {ex}");
                }
                finally {
                    // DTE Close() может завершиться без исключения, хотя пользователь отменил
                    // закрытие. Успех определяем по фактическому исчезновению модели из коллекции.
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
                // Даже при частичном успехе в историю попадают только реально закрытые вкладки.
                _closedTabsHistory.Push(closedEntries);
            }

            _virtualMenu.HideImmediately();
        }

        public void KeepOpen(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is TMEx.State.Document.TabItemDocument tabItemDocument) {
                // Сначала синхронизируем собственную модель групп, затем закрепляем документ
                // средствами VS, чтобы последующие shell-события уже увидели обычную вкладку.
                _tabCollectionManager.MovePreviewDocumentToDefaultGroup(tabItemDocument);
                tabItemDocument.ShellDocument.OpenDocumentAsPinned();
            }
        }

        public void CopyName(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.CopyTabValues(parameter, tabItem => tabItem.Caption, "name");
        }

        public void CopyPath(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.CopyTabValues(parameter, tabItem => tabItem.FullName, "path");
        }

        public void OpenLocation(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (parameter is TMEx.State.Document.TabItemDocument tabItemDocument) {
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

        private void CopyTabValues(
            object parameter,
            Func<TMEx.State.Document.TabItemBase, string> valueSelector,
            string valueKind
            ) {
            IReadOnlyList<TMEx.State.Document.TabItemBase> tabItems;
            if (parameter is IReadOnlyList<TMEx.State.Document.TabItemBase> selectedTabItems) {
                tabItems = selectedTabItems;
            }
            else if (parameter is TMEx.State.Document.TabItemBase anchorTabItem) {
                var selectedItems = _selectionCoordinator.SelectedItems;
                bool copySelection = selectedItems.Count > 1 &&
                    selectedItems.Any(entry => ReferenceEquals(entry.Item, anchorTabItem));
                tabItems = copySelection
                    ? selectedItems.Select(entry => entry.Item).ToList()
                    : new[] { anchorTabItem };
            }
            else {
                tabItems = _selectionCoordinator.SelectedItems.Select(entry => entry.Item).ToList();
            }

            if (tabItems.Count == 0) {
                return;
            }

            this.CopyText(string.Join(Environment.NewLine, tabItems.Select(valueSelector)), valueKind);
        }
    }
}
