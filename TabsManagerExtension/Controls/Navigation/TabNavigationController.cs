using System;
using System.Linq;
using System.Windows.Input;
using System.Windows.Threading;
using System.Collections.Generic;

using Helpers.Collections;
using TMEx = TabsManagerExtension;

namespace TabsManagerExtension.Controls.Navigation {
    // Определяет связь между выделением вкладок и активным VS-фреймом.
    internal enum TabSelectionActivationPolicy {
        // Активирует вкладку только простым кликом без Ctrl/Shift; мультивыбор не двигает frame.
        ActivateOnlyOnUnmodifiedPointerSelection,
        // Любой новый выбранный элемент становится активным, включая Space и SHIFT-диапазоны.
        ActivateLatestSelectedItem
    }


    internal interface ITabNavigationExtension {
        bool HandleKey(Key key, ModifierKeys modifiers);
    }


    internal sealed class TabNavigationController {
        private readonly GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase> _selectionCoordinator;
        private readonly Func<IReadOnlyList<TMEx.State.Document.TabItemBase>> _itemsProvider;
        private readonly List<ITabNavigationExtension> _extensions = new();
        private int _activationSuppressionDepth;

        private TabSelectionActivationPolicy _selectionActivationPolicy = TabSelectionActivationPolicy.ActivateOnlyOnUnmodifiedPointerSelection;
        public TabSelectionActivationPolicy SelectionActivationPolicy {
            get => _selectionActivationPolicy;
            set {
                _selectionActivationPolicy = value;
                // Политика активации одновременно задаёт семантику SHIFT-якоря, чтобы selection
                // и активный frame не расходились после последовательных диапазонных выделений.
                _selectionCoordinator.ShiftSelectionAnchorPolicy =
                    value == TabSelectionActivationPolicy.ActivateOnlyOnUnmodifiedPointerSelection
                        ? Helpers.Enums.ShiftSelectionAnchorPolicy.KeepInitialAnchor
                        : Helpers.Enums.ShiftSelectionAnchorPolicy.MoveToLatestRangeEndpoint;

                if (value == TabSelectionActivationPolicy.ActivateOnlyOnUnmodifiedPointerSelection) {
                    // В режиме сохранения frame каждый SHIFT-диапазон обязан начинаться
                    // от вкладки с фиолетовой рамкой, даже если внутренний anchor coordinator изменился.
                    _selectionCoordinator.ShiftSelectionAnchorProvider = () => {
                        var activeFrameItem = this.Items.FirstOrDefault(item => item.Metadata?.GetFlag("IsFrameActive") == true);
                        if (activeFrameItem == null) {
                            return SelectionAnchorResult<TMEx.State.Document.TabItemBase>.None;
                        }
                        return SelectionAnchorResult<TMEx.State.Document.TabItemBase>.FromItem(activeFrameItem);
                    };
                }
                else {
                    _selectionCoordinator.ShiftSelectionAnchorProvider = null;
                }
            }
        }

        public TabNavigationController(
            GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase> selectionCoordinator,
            Func<IReadOnlyList<TMEx.State.Document.TabItemBase>> itemsProvider
            ) {
            _selectionCoordinator = selectionCoordinator;
            _itemsProvider = itemsProvider;
            // Применяем не только значение по умолчанию, но и связанные настройки coordinator.
            this.SelectionActivationPolicy = _selectionActivationPolicy;
        }

        public TMEx.State.Document.TabItemBase? PrimaryItem => _selectionCoordinator.PrimarySelection?.Item;
        public IReadOnlyList<TMEx.State.Document.TabItemBase> Items => _itemsProvider();

        public void AddExtension(ITabNavigationExtension extension) {
            _extensions.Add(extension);
        }

        public bool HandleKey(Key key, ModifierKeys modifiers) {
            return _extensions.Any(extension => extension.HandleKey(key, modifiers));
        }

        public void OnPointerSelection(TMEx.State.Document.TabItemBase tabItem, ModifierKeys modifiers) {
            // Ctrl/Shift относятся к изменению набора выбранных вкладок. При политике
            // ActivateOnlyOnUnmodifiedPointerSelection они не должны менять открытый документ.
            bool hasSelectionModifier =
                (modifiers & ModifierKeys.Control) == ModifierKeys.Control ||
                (modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;

            bool shouldActivate = this.SelectionActivationPolicy == TabSelectionActivationPolicy.ActivateLatestSelectedItem ||
                !hasSelectionModifier;

            if (shouldActivate) {
                this.Activate(tabItem);
            }
        }

        public void ActivateLatestSelectionIfConfigured(TMEx.State.Document.TabItemBase tabItem) {
            // Клавиатурная навигация вызывает этот метод после изменения selection,
            // чтобы не дублировать знание о выбранной политике в каждом extension.
            if (this.SelectionActivationPolicy == TabSelectionActivationPolicy.ActivateLatestSelectedItem) {
                this.Activate(tabItem);
            }
        }

        public void SetSelectionWithoutActivation(TMEx.State.Document.TabItemBase tabItem, bool isSelected, ModifierKeys modifiers) {
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

        public void Activate(TMEx.State.Document.TabItemBase? tabItem) {
            if (tabItem is TMEx.State.Document.IActivatableTab activatableTab) {
                activatableTab.Activate();
            }
        }

        public void SelectAll() {
            foreach (var item in this.Items) {
                if (!item.IsSelected) {
                    this.SetSelectionWithoutActivation(item, true, ModifierKeys.Control);
                }
            }
        }
    }
}
