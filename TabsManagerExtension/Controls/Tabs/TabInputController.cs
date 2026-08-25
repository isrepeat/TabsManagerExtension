using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Collections.Generic;

using TabsManagerExtension.State.Document;


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
        private readonly Helpers.Collections.GroupsSelectionCoordinator<TabItemsGroupBase, TabItemBase> _selectionCoordinator;
        private readonly Navigation.TabNavigationController _navigationController;
        private readonly Navigation.KeyboardTabNavigationExtension _keyboardNavigation;
        private readonly ClosedTabsHistory _closedTabsHistory;
        private readonly TabAppearanceManager _appearanceManager;
        private readonly Action<IReadOnlyList<TabItemBase>> _closeTabs;
        private readonly Action _restoreClosedTabs;
        private readonly HashSet<TabItemControl> _tabItemControls = new();

        public TabInputController(
            FrameworkElement focusScope,
            UIElement focusTarget,
            ResourceDictionary rootResources,
            Dispatcher dispatcher,
            Func<bool> isEditMode,
            Func<double> scaleFactor,
            TabCollectionManager tabCollectionManager,
            Helpers.Collections.GroupsSelectionCoordinator<TabItemsGroupBase, TabItemBase> selectionCoordinator,
            Navigation.TabNavigationController navigationController,
            Navigation.KeyboardTabNavigationExtension keyboardNavigation,
            ClosedTabsHistory closedTabsHistory,
            TabAppearanceManager appearanceManager,
            Action<IReadOnlyList<TabItemBase>> closeTabs,
            Action restoreClosedTabs
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
            _closeTabs = closeTabs;
            _restoreClosedTabs = restoreClosedTabs;
        }

        public void Register(TabItemControl control) {
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

        public bool HandleEditKey(Key key, ModifierKeys modifiers) {
            if (!_isEditMode()) {
                return false;
            }

            bool isControlPressed = (modifiers & ModifierKeys.Control) == ModifierKeys.Control;
            if (key == Key.F2 && modifiers == ModifierKeys.None) {
                this.FindControl(_keyboardNavigation.FocusedItem)?.BeginRename();
                return true;
            }
            if (key == Key.Escape && modifiers == ModifierKeys.None) {
                var focusedItem = _keyboardNavigation.FocusedItem;
                if (focusedItem != null) {
                    _selectionCoordinator.SetSelection(focusedItem, true, ModifierKeys.None);
                    _keyboardNavigation.RestoreInputTarget();
                }

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
                var itemsToClose = focusedItem.IsSelected && selectedItems.Count > 1
                    ? selectedItems.Select(entry => entry.Item).ToList()
                    : new List<TabItemBase> { focusedItem };
                _closeTabs(itemsToClose);

                var nextFocusedItem = _tabCollectionManager.AllTabs.FirstOrDefault();
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

                _restoreClosedTabs();
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

                return true;
            }

            return key != Key.Z || _closedTabsHistory.Count > 0;
        }

        public void HandlePointerNavigation(TabItemControl source, ModifierKeys modifiers) {
            if (source.DataContext is not TabItemBase tabItem) {
                return;
            }

            _navigationController.OnPointerSelection(tabItem, modifiers);
            if (_isEditMode()) {
                _keyboardNavigation.FocusItem(tabItem);
                _keyboardNavigation.RestoreInputTarget();
            }
        }

        public void HandleFocusedItemChanged(TabItemBase? tabItem) {
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
            _dispatcher.BeginInvoke(new Action(this.FocusInputTarget), DispatcherPriority.ContextIdle);
        }

        public void FocusInputTarget() {
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

        private TabItemControl? FindControl(TabItemBase? tabItem) {
            return tabItem == null
                ? null
                : _tabItemControls.FirstOrDefault(candidate => ReferenceEquals(candidate.DataContext, tabItem));
        }
    }
}
