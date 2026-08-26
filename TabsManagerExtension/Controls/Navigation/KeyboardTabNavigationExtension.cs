using System;
using System.Linq;
using System.Windows.Input;

using TMEx = TabsManagerExtension;

namespace TabsManagerExtension.Controls.Navigation {
    internal sealed class KeyboardTabNavigationExtension : ITabNavigationExtension {
        private readonly TabNavigationController _navigationController;

        public event Action<TMEx.State.Document.TabItemBase?>? FocusedItemChanged;
        public event Action? InputTargetRestoreRequested;

        public TMEx.State.Document.TabItemBase? FocusedItem { get; private set; }

        public KeyboardTabNavigationExtension(TabNavigationController navigationController) {
            _navigationController = navigationController;
        }

        public void InitializeFocus() {
            this.SetFocusedItem(_navigationController.PrimaryItem ?? _navigationController.Items.FirstOrDefault());
            this.InputTargetRestoreRequested?.Invoke();
        }

        public void FocusItem(TMEx.State.Document.TabItemBase tabItem) {
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
                // Space меняет только membership текущей вкладки в мультивыборе. Активный
                // VS-фрейм остаётся прежним, если выбранная policy не требует обратного.
                _navigationController.SetSelectionWithoutActivation(
                    this.FocusedItem,
                    !this.FocusedItem.IsSelected,
                    ModifierKeys.Control
                );

                if (this.FocusedItem.IsSelected) {
                    _navigationController.ActivateLatestSelectionIfConfigured(this.FocusedItem);
                }

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
                // Shift+стрелка одновременно перемещает навигационный фокус и расширяет диапазон
                // от anchor, который определяется текущей политикой активации.
                _navigationController.SetSelectionWithoutActivation(targetItem, true, ModifierKeys.Shift);
                _navigationController.ActivateLatestSelectionIfConfigured(targetItem);
            }

            return true;
        }

        private void SetFocusedItem(TMEx.State.Document.TabItemBase? tabItem) {
            if (!ReferenceEquals(this.FocusedItem, tabItem)) {
                this.FocusedItem = tabItem;
                this.FocusedItemChanged?.Invoke(tabItem);
            }
        }
    }
}
