using System;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Collections.ObjectModel;

using Microsoft.VisualStudio.Shell;

using Helpers.Ex;
using TabsManagerExtension.State.Document;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Синхронизирует виртуальное hover-меню с текущей вкладкой и project context.</summary>
    internal sealed class VirtualMenuController {
        private readonly VirtualMenuControl _virtualMenuControl;
        private readonly ObservableCollection<Helpers.IMenuItem> _menuItems;
        private readonly TabMenuItemFactory _menuItemFactory;
        private readonly ProjectContextController _projectContextController;

        public VirtualMenuController(
            VirtualMenuControl virtualMenuControl,
            ObservableCollection<Helpers.IMenuItem> menuItems,
            TabMenuItemFactory menuItemFactory,
            ProjectContextController projectContextController
            ) {
            _virtualMenuControl = virtualMenuControl;
            _menuItems = menuItems;
            _menuItemFactory = menuItemFactory;
            _projectContextController = projectContextController;
        }

        public void HandleCloseButtonMouseEnter(object sender) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (sender is not Button closeButton) {
                return;
            }

            var tabItemControl = Helpers.VisualTree.FindParentByType<TabItemControl>(closeButton);
            var listViewItem = tabItemControl == null
                ? null
                : Helpers.VisualTree.FindParentByType<ListViewItem>(tabItemControl);
            if (tabItemControl == null || listViewItem?.DataContext is not TabItemDocument tabItemDocument) {
                return;
            }

            this.ClearPreviousMenuFlag();
            Point screenPoint = tabItemControl.ex_ToDpiAwareScreen(new Point(tabItemControl.ActualWidth + 20, -60));
            _virtualMenuControl.Show(screenPoint, tabItemDocument);
        }

        public void HandleInteractiveAreaMouseEnter(object sender) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!_virtualMenuControl.IsMenuOpen ||
                sender is not TabItemControl tabItemControl ||
                tabItemControl.DataContext is not TabItemDocument tabItemDocument ||
                ReferenceEquals(_virtualMenuControl.CurrentMenuDataContext, tabItemDocument)) {

                return;
            }

            this.ClearPreviousMenuFlag();
            Point screenPoint = tabItemControl.ex_ToDpiAwareScreen(new Point(tabItemControl.ActualWidth + 20, -60));
            _virtualMenuControl.Show(screenPoint, tabItemDocument);
        }

        public void HandleInteractiveAreaMouseLeave() {
            ThreadHelper.ThrowIfNotOnUIThread();
            _virtualMenuControl.Hide();
        }

        public void Open(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _virtualMenuControl.HideChild();

            if (parameter is not MenuControl.MenuOpeningArgs openingArgs ||
                openingArgs.DataContext is not TabItemBase tabItem) {

                return;
            }
            if (tabItem is TabItemWindow) {
                openingArgs.ShouldOpen = false;
                return;
            }
            if (tabItem is not TabItemDocument tabItemDocument) {
                return;
            }

            tabItemDocument.Metadata?.SetFlag("IsVirtualMenuOpenned", true);
            var newItems = _menuItemFactory.CreateVirtualMenu(tabItemDocument);
            if (tabItemDocument.ShellDocument != null) {
                _projectContextController.AppendProjectMenuItems(
                    newItems,
                    tabItemDocument,
                    _menuItemFactory
                );
                this.UpdateItems(newItems);
            }
        }

        public void Close(object parameter) {
            _virtualMenuControl.HideChild();

            if (parameter is MenuControl.MenuClosedArgs closedArgs && closedArgs.DataContext is TabItemBase tabItem) {
                tabItem.Metadata?.SetFlag("IsVirtualMenuOpenned", false);
            }
        }

        private void ClearPreviousMenuFlag() {
            if (_virtualMenuControl.CurrentMenuDataContext is TabItemDocument previousTabItemDocument) {
                previousTabItemDocument.Metadata?.SetFlag("IsVirtualMenuOpenned", false);
            }
        }

        private void UpdateItems(IReadOnlyList<Helpers.IMenuItem> newItems) {
            int commonCount = Math.Min(_menuItems.Count, newItems.Count);
            for (int index = 0; index < commonCount; index++) {
                var currentItem = _menuItems[index];
                var newItem = newItems[index];
                if (currentItem is Helpers.MenuItemHeader currentHeader &&
                    newItem is Helpers.MenuItemHeader newHeader) {

                    currentHeader.Header = newHeader.Header;
                    continue;
                }
                if (currentItem is Helpers.MenuItemSeparator && newItem is Helpers.MenuItemSeparator) {
                    continue;
                }
                if (currentItem is Helpers.MenuItemCommand currentCommand &&
                    newItem is Helpers.MenuItemCommand newCommand &&
                    GetItemKind(currentCommand) == GetItemKind(newCommand)) {

                    currentCommand.Header = newCommand.Header;
                    currentCommand.CommandParameterContext = newCommand.CommandParameterContext;
                    continue;
                }

                _menuItems[index] = newItem;
            }

            while (_menuItems.Count > newItems.Count) {
                _menuItems.RemoveAt(_menuItems.Count - 1);
            }
            for (int index = commonCount; index < newItems.Count; index++) {
                _menuItems.Add(newItems[index]);
            }
        }

        private static Type GetItemKind(Helpers.MenuItemCommand menuItem) {
            return menuItem.CommandParameterContext switch {
                DocumentProjectReferencesInfo.RefEntry => typeof(DocumentProjectReferencesInfo.RefEntry),
                DocumentProjectReferencesInfo.GroupContextEntry => typeof(DocumentProjectReferencesInfo.GroupContextEntry),
                DocumentProjectReferencesInfo => typeof(DocumentProjectReferencesInfo),
                _ => typeof(TabItemBase)
            };
        }
    }
}
