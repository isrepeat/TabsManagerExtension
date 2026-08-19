using System;
using System.Linq;
using System.Windows.Input;
using TabsManagerExtension.State.Document;


namespace TabsManagerExtension.Controls.Navigation {
    internal sealed class KeyboardTabNavigationExtension : ITabNavigationExtension {
        private readonly TabNavigationController _navigationController;

        public event Action<TabItemBase?>? FocusedItemChanged;
        public event Action? InputTargetRestoreRequested;

        public TabItemBase? FocusedItem { get; private set; }

        public KeyboardTabNavigationExtension(TabNavigationController navigationController) {
            _navigationController = navigationController;
        }

        public void InitializeFocus() {
            this.SetFocusedItem(_navigationController.PrimaryItem ?? _navigationController.Items.FirstOrDefault());
            this.InputTargetRestoreRequested?.Invoke();
        }

        public void FocusItem(TabItemBase tabItem) {
            this.SetFocusedItem(tabItem);
        }

        public void RestoreInputTarget() {
            this.InputTargetRestoreRequested?.Invoke();
        }

        public void ClearFocus() {
            this.SetFocusedItem(null);
        }

        public bool HandleKey(Key key, ModifierKeys modifiers) {
            if (this.FocusedItem == null) {
                return false;
            }

            if (key == Key.Enter) {
                _navigationController.Activate(this.FocusedItem);
                this.InputTargetRestoreRequested?.Invoke();
                return true;
            }

            if (key == Key.Space) {
                _navigationController.SetSelectionWithoutActivation(
                    this.FocusedItem,
                    !this.FocusedItem.IsSelected,
                    ModifierKeys.Control
                );

                this.InputTargetRestoreRequested?.Invoke();
                return true;
            }

            if (key != Key.Up && key != Key.Down) {
                return false;
            }

            var items = _navigationController.Items;
            int currentIndex = items.ToList().FindIndex(item => ReferenceEquals(item, this.FocusedItem));
            int targetIndex = key == Key.Up ? currentIndex - 1 : currentIndex + 1;
            if (currentIndex < 0 || targetIndex < 0 || targetIndex >= items.Count) {
                return true;
            }

            var targetItem = items[targetIndex];
            this.SetFocusedItem(targetItem);

            if ((modifiers & ModifierKeys.Shift) == ModifierKeys.Shift) {
                _navigationController.SetSelectionWithoutActivation(targetItem, true, ModifierKeys.Shift);
            }

            return true;
        }

        private void SetFocusedItem(TabItemBase? tabItem) {
            if (!ReferenceEquals(this.FocusedItem, tabItem)) {
                this.FocusedItem = tabItem;
                this.FocusedItemChanged?.Invoke(tabItem);
            }
        }
    }
}
