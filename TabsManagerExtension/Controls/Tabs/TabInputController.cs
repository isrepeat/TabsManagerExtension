using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Collections.Generic;

using TMEx = TabsManagerExtension;

namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Управляет клавиатурной и мышиной навигацией, edit mode и WPF-фокусом вкладок.</summary>
    internal sealed class TabInputController {
        private readonly FrameworkElement _focusScope;
        private readonly UIElement _focusTarget;
        private readonly ResourceDictionary _rootResources;
        private readonly Dispatcher _dispatcher;
        private readonly Func<bool> _isEditMode;
        private readonly Func<double> _scaleFactor;
        private readonly TabCollectionManager _tabCollectionManager;
        private readonly Helpers.Collections.GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase> _selectionCoordinator;
        private readonly Navigation.TabNavigationController _navigationController;
        private readonly Navigation.KeyboardTabNavigationExtension _keyboardNavigation;
        private readonly ClosedTabsHistory _closedTabsHistory;
        private readonly TabAppearanceManager _appearanceManager;
        private readonly Action<IReadOnlyList<TMEx.State.Document.TabItemBase>> _onCloseTabs;
        private readonly Action _onRestoreClosedTabs;
        private readonly Action _onCopySelectedTabNames;
        private readonly HashSet<TabItemControl> _tabItemControls = new();

        public TabInputController(
            FrameworkElement focusScope,
            UIElement focusTarget,
            ResourceDictionary rootResources,
            Dispatcher dispatcher,
            Func<bool> isEditMode,
            Func<double> scaleFactor,
            TabCollectionManager tabCollectionManager,
            Helpers.Collections.GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase> selectionCoordinator,
            Navigation.TabNavigationController navigationController,
            Navigation.KeyboardTabNavigationExtension keyboardNavigation,
            ClosedTabsHistory closedTabsHistory,
            TabAppearanceManager appearanceManager,
            Action<IReadOnlyList<TMEx.State.Document.TabItemBase>> onCloseTabs,
            Action onRestoreClosedTabs,
            Action onCopySelectedTabNames
            ) {
            _focusScope = focusScope;
            _focusTarget = focusTarget;
            _rootResources = rootResources;
            _dispatcher = dispatcher;
            _isEditMode = isEditMode;
            _scaleFactor = scaleFactor;
            _tabCollectionManager = tabCollectionManager;
            _selectionCoordinator = selectionCoordinator;
            _navigationController = navigationController;
            _keyboardNavigation = keyboardNavigation;
            _closedTabsHistory = closedTabsHistory;
            _appearanceManager = appearanceManager;
            _onCloseTabs = onCloseTabs;
            _onRestoreClosedTabs = onRestoreClosedTabs;
            _onCopySelectedTabNames = onCopySelectedTabNames;
        }

        public void Register(TabItemControl control) {
            // WPF может пересоздавать визуальные контейнеры при сортировке и виртуализации.
            // Новый экземпляр сразу получает актуальные scale, appearance и edit-focus.
            _tabItemControls.Add(control);
            _appearanceManager.ApplyScale(control, _scaleFactor());
            _appearanceManager.ApplyAppearance(control);
            if (ReferenceEquals(control.DataContext, _keyboardNavigation.FocusedItem)) {
                control.IsEditFocused = true;
            }
        }

        public void Unregister(TabItemControl control) {
            _tabItemControls.Remove(control);
        }

        // Обрабатывает команды режима редактирования вкладок и сообщает, должна ли клавиша
        // считаться обработанной до передачи в стандартный редактор Visual Studio.
        public bool HandleEditKey(Key key, ModifierKeys modifiers) {
            if (!_isEditMode()) {
                return false;
            }

            bool isControlPressed = (modifiers & ModifierKeys.Control) == ModifierKeys.Control;
            if (key == Key.F2 && modifiers == ModifierKeys.None) {
                var focusedItem = _keyboardNavigation.FocusedItem;
                if (focusedItem == null) {
                    return true;
                }

                var focusedControl = this.FindControl(focusedItem);
                if (focusedControl?.IsRenaming == true) {
                    focusedControl.ToggleRenameSelection();
                    return true;
                }

                var renameGroupTabItems = focusedItem is TMEx.State.Document.TabItemDocument focusedDocumentTabItem
                    ? TabRenameService.GetSelectedRenameGroup(
                        focusedDocumentTabItem,
                        _selectionCoordinator.SelectedItems.Select(entry => entry.Item)
                    )
                    : Array.Empty<TMEx.State.Document.TabItemDocument>();

                if (renameGroupTabItems.Count == 0) {
                    // Обычное F2 редактирует только navigation focus. Явно сворачиваем selection,
                    // чтобы интерфейс не создавал впечатления пакетного переименования.
                    _selectionCoordinator.SetSelection(focusedItem, true, ModifierKeys.None);
                }

                focusedControl?.BeginRename(renameGroupTabItems);
                return true;
            }
            if (key == Key.Escape && modifiers == ModifierKeys.None) {
                // Escape сворачивает мультивыбор к навигационно сфокусированной вкладке,
                // после чего возвращает скрытую WPF-цель ввода для следующих команд.
                var focusedItem = _keyboardNavigation.FocusedItem;
                if (focusedItem != null) {
                    _selectionCoordinator.SetSelection(focusedItem, true, ModifierKeys.None);
                    _keyboardNavigation.RestoreInputTarget();
                }

                return true;
            }
            if (isControlPressed && key == Key.C) {
                _onCopySelectedTabNames();
                return true;
            }
            if (isControlPressed && key == Key.A) {
                _navigationController.SelectAll();
                _keyboardNavigation.RestoreInputTarget();
                return true;
            }
            if (key == Key.Delete && modifiers == ModifierKeys.None) {
                var focusedItem = _keyboardNavigation.FocusedItem;
                if (focusedItem == null) {
                    return true;
                }

                var selectedItems = _selectionCoordinator.SelectedItems;
                // Delete следует той же семантике, что и контекстное Close: если focused item
                // входит в мультивыбор, закрывается selection, иначе только focused item.
                var itemsToClose = focusedItem.IsSelected && selectedItems.Count > 1
                    ? selectedItems.Select(entry => entry.Item).ToList()
                    : new List<TMEx.State.Document.TabItemBase> { focusedItem };
                _onCloseTabs(itemsToClose);

                var nextFocusedItem = _tabCollectionManager.AllTabs.FirstOrDefault();
                // Закрытие удаляет прежний FocusedItem вместе с его visual container. Назначаем
                // новый навигационный фокус явно, чтобы клавиатурный режим не потерял target.
                if (nextFocusedItem != null) {
                    _keyboardNavigation.FocusItem(nextFocusedItem);
                    _keyboardNavigation.RestoreInputTarget();
                }

                return true;
            }
            if (isControlPressed && key == Key.Z) {
                if (_closedTabsHistory.Count == 0) {
                    return false;
                }

                _onRestoreClosedTabs();
                return true;
            }

            return _navigationController.HandleKey(key, modifiers);
        }

        public bool CanHandleRedirectedKey(Key key) {
            if (!_focusScope.IsKeyboardFocusWithin) {
                return false;
            }
            if (Keyboard.FocusedElement is DependencyObject focusedElement &&
                Helpers.VisualTree.FindParentByType<TabItemControl>(focusedElement)?.IsRenaming == true) {

                // Во время inline rename command filter должен оставить клавишу внутри панели:
                // окончательное решение о Enter/Escape принимает TextBox вкладки.
                return true;
            }

            // Не перехватываем Ctrl+Z у редактора, когда восстанавливать больше нечего.
            return key != Key.Z || _closedTabsHistory.Count > 0;
        }

        public void HandlePointerNavigation(TabItemControl source, ModifierKeys modifiers) {
            if (source.DataContext is not TMEx.State.Document.TabItemBase tabItem) {
                return;
            }

            _navigationController.OnPointerSelection(tabItem, modifiers);
            if (_isEditMode()) {
                // Обычный клик обновляет selection, но в edit mode дополнительно переносим
                // отдельный навигационный фокус, по которому работают F2/Delete/стрелки.
                _keyboardNavigation.FocusItem(tabItem);
                _keyboardNavigation.RestoreInputTarget();
            }
        }

        public void HandleFocusedItemChanged(TMEx.State.Document.TabItemBase? tabItem) {
            // IsEditFocused — визуальный маркер ровно одного контейнера; он не совпадает
            // с IsSelected при множественном выборе.
            var previousControl = _tabItemControls.FirstOrDefault(control => control.IsEditFocused);
            if (previousControl != null) {
                previousControl.IsEditFocused = false;
            }

            var control = this.FindControl(tabItem);
            if (control == null) {
                return;
            }

            control.IsEditFocused = true;
            control.BringIntoView();
            this.FocusInputTarget();
        }

        public void RestoreInputTarget() {
            // Откладываем возврат фокуса до ContextIdle, чтобы завершающийся mouse/key event
            // или закрытие контейнера не перехватили фокус обратно.
            VsixThreadHelper.RunOnUiThread(
                _dispatcher,
                this.FocusInputTarget,
                DispatcherPriority.ContextIdle
            );
        }

        public void FocusInputTarget() {
            // Logical focus, UIElement focus и keyboard focus устанавливаются вместе: одной
            // операции недостаточно для стабильной маршрутизации WPF и VS command filter.
            FocusManager.SetFocusedElement(_focusScope, _focusTarget);
            _focusTarget.Focus();
            Keyboard.Focus(_focusTarget);
        }

        public void ApplyScale() {
            foreach (var control in _tabItemControls) {
                _appearanceManager.ApplyScale(control, _scaleFactor());
            }
        }

        public void ApplyAppearance() {
            _appearanceManager.ApplyAppearance(_rootResources, _tabItemControls);
        }

        private TabItemControl? FindControl(TMEx.State.Document.TabItemBase? tabItem) {
            // Ищем по ссылке на view-model: Caption и FullName могут измениться при rename,
            // а один и тот же путь временно встречается в пересоздаваемых контейнерах.
            return tabItem == null
                ? null
                : _tabItemControls.FirstOrDefault(candidate => ReferenceEquals(candidate.DataContext, tabItem));
        }
    }
}
