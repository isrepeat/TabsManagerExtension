using System;
using System.Linq;
using System.Windows.Input;
using System.Windows.Threading;
using System.Collections.Generic;
using Helpers.Collections;
using TabsManagerExtension.State.Document;


namespace TabsManagerExtension.Controls.Navigation {
    internal interface ITabNavigationExtension {
        bool HandleKey(Key key, ModifierKeys modifiers);
    }


    internal sealed class TabNavigationController {
        private readonly GroupsSelectionCoordinator<TabItemsGroupBase, TabItemBase> _selectionCoordinator;
        private readonly Func<IReadOnlyList<TabItemBase>> _itemsProvider;
        private readonly List<ITabNavigationExtension> _extensions = new();
        private int _activationSuppressionDepth;

        public TabNavigationController(
            GroupsSelectionCoordinator<TabItemsGroupBase, TabItemBase> selectionCoordinator,
            Func<IReadOnlyList<TabItemBase>> itemsProvider
            ) {
            _selectionCoordinator = selectionCoordinator;
            _itemsProvider = itemsProvider;
        }

        public TabItemBase? PrimaryItem => _selectionCoordinator.PrimarySelection?.Item;
        public IReadOnlyList<TabItemBase> Items => _itemsProvider();

        public void AddExtension(ITabNavigationExtension extension) {
            _extensions.Add(extension);
        }

        public bool HandleKey(Key key, ModifierKeys modifiers) {
            return _extensions.Any(extension => extension.HandleKey(key, modifiers));
        }

        public void OnItemSelectionChanged(TabItemBase tabItem, bool isSelected, bool isActivatedExternally) {
            if (isSelected && !isActivatedExternally && _activationSuppressionDepth == 0) {
                this.Activate(tabItem);
            }
        }

        public void OnSelectionStateChanged(Helpers.Enums.SelectionState selectionState) {
            if (selectionState == Helpers.Enums.SelectionState.Single && _activationSuppressionDepth == 0) {
                this.Activate(this.PrimaryItem);
            }
        }

        public void SetSelectionWithoutActivation(TabItemBase tabItem, bool isSelected, ModifierKeys modifiers) {
            _activationSuppressionDepth++;
            try {
                _selectionCoordinator.SetSelection(tabItem, isSelected, modifiers);
            }
            catch {
                _activationSuppressionDepth--;
                throw;
            }

            // Coordinator рассылает изменения selection через Dispatcher. Сохраняем запрет
            // активации до завершения этой рассылки, иначе Space/Shift+Arrow активируют frame.
            _ = Dispatcher.CurrentDispatcher.BeginInvoke(new Action(() => _activationSuppressionDepth--));
        }

        public void Activate(TabItemBase? tabItem) {
            if (tabItem is IActivatableTab activatableTab) {
                activatableTab.Activate();
            }
        }
    }
}
