using System;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.VisualStudio.Shell;

using Helpers.Ex;
using TMEx = TabsManagerExtension;

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
            // Кнопка вложена в template вкладки, а ListViewItem хранит её data context.
            // Поднимаемся по visual tree вместо зависимости от конкретной структуры шаблона.
            var listViewItem = tabItemControl == null
                ? null
                : Helpers.VisualTree.FindParentByType<ListViewItem>(tabItemControl);
            if (tabItemControl == null || listViewItem?.DataContext is not TMEx.State.Document.TabItemDocument tabItemDocument) {
                return;
            }

            this.ClearPreviousMenuFlag();
            // VirtualMenuControl принимает экранные координаты; helper учитывает DPI монитора,
            // на котором в данный момент находится tool window.
            Point screenPoint = tabItemControl.ex_ToDpiAwareScreen(new Point(tabItemControl.ActualWidth + 20, -60));
            _virtualMenuControl.Show(screenPoint, tabItemDocument);
        }

        public void HandleInteractiveAreaMouseEnter(object sender) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!_virtualMenuControl.IsMenuOpen ||
                sender is not TabItemControl tabItemControl ||
                tabItemControl.DataContext is not TMEx.State.Document.TabItemDocument tabItemDocument) {

                return;
            }

            if (!ReferenceEquals(_virtualMenuControl.CurrentMenuDataContext, tabItemDocument)) {
                this.ClearPreviousMenuFlag();
            }

            // Уже открытое hover-меню переиспользуется для соседней вкладки без промежуточного
            // закрытия popup, иначе при движении мыши появляется заметное мерцание. Повторный
            // вход на ту же вкладку тоже вызывает Show: он отменяет отложенное скрытие меню.
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
                openingArgs.DataContext is not TMEx.State.Document.TabItemBase tabItem) {

                return;
            }
            if (tabItem is TMEx.State.Document.TabItemWindow) {
                // Project context применим только к документам. Для tool window полностью
                // отменяем открытие, а не показываем пустое меню.
                openingArgs.ShouldOpen = false;
                return;
            }
            if (tabItem is not TMEx.State.Document.TabItemDocument tabItemDocument) {
                return;
            }

            tabItemDocument.Metadata?.SetFlag("IsVirtualMenuOpenned", true);
            // Сначала создаются стабильные общие команды, затем ProjectContextController
            // дописывает динамическую часть, зависящую от текущего solution graph.
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

            if (parameter is MenuControl.MenuClosedArgs closedArgs && closedArgs.DataContext is TMEx.State.Document.TabItemBase tabItem) {
                tabItem.Metadata?.SetFlag("IsVirtualMenuOpenned", false);
            }
        }

        private void ClearPreviousMenuFlag() {
            // При переходе между вкладками событие Close для предыдущего data context может
            // не успеть прийти, поэтому снимаем визуальный metadata-флаг вручную.
            if (_virtualMenuControl.CurrentMenuDataContext is TMEx.State.Document.TabItemDocument previousTabItemDocument) {
                previousTabItemDocument.Metadata?.SetFlag("IsVirtualMenuOpenned", false);
            }
        }

        private void UpdateItems(IReadOnlyList<Helpers.IMenuItem> newItems) {
            // По возможности обновляем существующие объекты на месте. Полная замена коллекции
            // пересоздаёт visual tree popup и сбрасывает hover/child-menu во время наведения.
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

                    // Команду можно переиспользовать только для того же вида context parameter:
                    // обработчики project/reference/group entries интерпретируют его по-разному.
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
                TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry => typeof(TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry),
                TMEx.State.Document.DocumentProjectReferencesInfo.GroupContextEntry => typeof(TMEx.State.Document.DocumentProjectReferencesInfo.GroupContextEntry),
                TMEx.State.Document.DocumentProjectReferencesInfo => typeof(TMEx.State.Document.DocumentProjectReferencesInfo),
                _ => typeof(TMEx.State.Document.TabItemBase)
            };
        }
    }
}
