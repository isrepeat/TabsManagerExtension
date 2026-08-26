using System;
using System.IO;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using Helpers.Ex;

namespace TabsManagerExtension.Controls {
    public partial class TabItemControl : Helpers.BaseUserControl {
        public string Title {
            get { return (string)this.GetValue(TitleProperty); }
            set { this.SetValue(TitleProperty, value); }
        }
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(
                nameof(Title),
                typeof(string),
                typeof(TabItemControl),
                new PropertyMetadata("Document Title"));


        public bool IsSelected {
            get { return (bool)this.GetValue(IsSelectedProperty); }
            set { this.SetValue(IsSelectedProperty, value); }
        }
        public static readonly DependencyProperty IsSelectedProperty =
            DependencyProperty.Register(
                nameof(IsSelected),
                typeof(bool),
                typeof(TabItemControl),
                new PropertyMetadata(false));


        public bool IsEditMode {
            get { return (bool)this.GetValue(IsEditModeProperty); }
            set { this.SetValue(IsEditModeProperty, value); }
        }
        public static readonly DependencyProperty IsEditModeProperty =
            DependencyProperty.Register(
                nameof(IsEditMode),
                typeof(bool),
                typeof(TabItemControl),
                new PropertyMetadata(false));


        public bool IsMultipleSelection {
            get { return (bool)this.GetValue(IsMultipleSelectionProperty); }
            set { this.SetValue(IsMultipleSelectionProperty, value); }
        }
        public static readonly DependencyProperty IsMultipleSelectionProperty =
            DependencyProperty.Register(
                nameof(IsMultipleSelection),
                typeof(bool),
                typeof(TabItemControl),
                new PropertyMetadata(false));


        public bool IsEditFocused {
            get { return (bool)this.GetValue(IsEditFocusedProperty); }
            set { this.SetValue(IsEditFocusedProperty, value); }
        }
        public static readonly DependencyProperty IsEditFocusedProperty =
            DependencyProperty.Register(
                nameof(IsEditFocused),
                typeof(bool),
                typeof(TabItemControl),
                new PropertyMetadata(false));


        public DataTemplate ControlPanelPrimarySlotTemplate {
            get => (DataTemplate)this.GetValue(ControlPanelPrimarySlotTemplateProperty);
            set => this.SetValue(ControlPanelPrimarySlotTemplateProperty, value);
        }
        public static readonly DependencyProperty ControlPanelPrimarySlotTemplateProperty =
            DependencyProperty.Register(nameof(ControlPanelPrimarySlotTemplate),
                typeof(DataTemplate),
                typeof(TabItemControl),
                new PropertyMetadata(null));


        public DataTemplate ControlPanelSecondarySlotTemplate {
            get => (DataTemplate)this.GetValue(ControlPanelSecondarySlotTemplateProperty);
            set => this.SetValue(ControlPanelSecondarySlotTemplateProperty, value);
        }
        public static readonly DependencyProperty ControlPanelSecondarySlotTemplateProperty =
            DependencyProperty.Register(nameof(ControlPanelSecondarySlotTemplate),
                typeof(DataTemplate),
                typeof(TabItemControl),
                new PropertyMetadata(null));


        public DataTemplate ContextMenuTemplate {
            get => (DataTemplate)this.GetValue(ContextMenuTemplateProperty);
            set => this.SetValue(ContextMenuTemplateProperty, value);
        }
        public static readonly DependencyProperty ContextMenuTemplateProperty =
            DependencyProperty.Register(nameof(ContextMenuTemplate),
                typeof(DataTemplate),
                typeof(TabItemControl),
                new PropertyMetadata(null));


        private bool _isMouseInside = false;
        public bool IsMouseInside {
            get { return _isMouseInside; }
            set {
                if (_isMouseInside != value) {
                    _isMouseInside = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _isRenaming;
        public bool IsRenaming {
            get => _isRenaming;
            private set {
                if (_isRenaming != value) {
                    _isRenaming = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _renameText = string.Empty;
        public string RenameText {
            get => _renameText;
            set {
                if (_renameText != value) {
                    _renameText = value;
                    OnPropertyChanged();
                }
            }
        }


        private WeakReference<MenuControl>? _cachedWeakMenuControl;
        private IReadOnlyList<State.Document.TabItemDocument> _renameGroupTabItems = Array.Empty<State.Document.TabItemDocument>();
        private bool _isCompletingRename;
        private bool _isRenameFileNameOnlySelected;
        private DateTime _lastPointerNavigationUtc;
        private bool _restoreRenameFocusAfterPointerActivation;

        public TabItemControl() {
            this.InitializeComponent();
            this.Loaded += this.OnLoaded;
            this.Unloaded += this.OnUnloaded;
            this.PreviewKeyDown += this.OnPreviewKeyDown;
            // Внутренний CheckBox завершает MouseLeftButtonUp как handled. Подписка с
            // handledEventsToo нужна для общей обработки выбора и активации вкладки.
            this.AddHandler(MouseLeftButtonUpEvent, new MouseButtonEventHandler(this.OnMouseLeftButtonUp), true);
            this.MouseEnter += this.OnMouseEnter;
            this.MouseLeave += this.OnMouseLeave;
            this.MouseRightButtonUp += this.OnMouseRightButtonUpHandler;
        }

        private void OnLoaded(object sender, RoutedEventArgs e) {
            this.FindAndCacheMenuControl();
            Helpers.VisualTree.FindParentByType<TabsManagerToolWindowControl>(this)?.RegisterTabItemControl(this);
        }

        private void OnUnloaded(object sender, RoutedEventArgs e) {
            Helpers.VisualTree.FindParentByType<TabsManagerToolWindowControl>(this)?.UnregisterTabItemControl(this);
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e) {
            if (!this.IsEditMode || this.IsRenaming) {
                return;
            }

            var owner = Helpers.VisualTree.FindParentByType<TabsManagerToolWindowControl>(this);
            if (owner?.HandleTabEditKey(e.Key, Keyboard.Modifiers) == true) {
                e.Handled = true;
            }
        }

        public void BeginRename(IReadOnlyList<State.Document.TabItemDocument> renameGroupTabItems) {
            this.RenameText = this.Title;
            _renameGroupTabItems = renameGroupTabItems;
            _restoreRenameFocusAfterPointerActivation = DateTime.UtcNow - _lastPointerNavigationUtc <= TimeSpan.FromMilliseconds(500);
            this.IsRenaming = true;
            this.Dispatcher.BeginInvoke(new Action(() => {
                this.RenameTextBox.Focus();
                Keyboard.Focus(this.RenameTextBox);
                this.SelectRenameFileName();
            }), DispatcherPriority.Input);
        }

        public void ToggleRenameSelection() {
            if (!this.IsRenaming) {
                return;
            }

            if (_isRenameFileNameOnlySelected) {
                this.RenameTextBox.SelectAll();
                _isRenameFileNameOnlySelected = false;
                return;
            }

            this.SelectRenameFileName();
        }

        private void SelectRenameFileName() {
            string text = this.RenameTextBox.Text;
            string extension = Path.GetExtension(text);
            this.RenameTextBox.Select(0, text.Length - extension.Length);
            _isRenameFileNameOnlySelected = true;
        }

        private void RenameTextBox_OnKeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.F2 && Keyboard.Modifiers == ModifierKeys.None) {
                this.ToggleRenameSelection();
                e.Handled = true;
            }
            else if (e.Key == Key.Home || e.Key == Key.End) {
                // Visual Studio может передать Home/End как editor-команду, поэтому явно
                // снимаем выделение и ставим каретку в край.
                int targetPosition = e.Key == Key.Home ? 0 : this.RenameTextBox.Text.Length;
                bool extendSelection = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
                if (extendSelection) {
                    int anchorPosition = this.RenameTextBox.CaretIndex;
                    this.RenameTextBox.Select(
                        Math.Min(anchorPosition, targetPosition),
                        Math.Abs(anchorPosition - targetPosition)
                    );
                }
                else {
                    this.RenameTextBox.Select(targetPosition, 0);
                }

                e.Handled = true;
            }
            else if (e.Key == Key.Enter) {
                this.CompleteRename();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape) {
                this.IsRenaming = false;
                _renameGroupTabItems = Array.Empty<State.Document.TabItemDocument>();
                _restoreRenameFocusAfterPointerActivation = false;
                this.RestoreOwnerInputTarget();
                e.Handled = true;
            }
        }

        private void RenameTextBox_OnLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e) {
            if (_restoreRenameFocusAfterPointerActivation) {
                // Простая активация вкладки в VS завершается асинхронно. Если F2 был нажат
                // сразу после клика, поздняя активация фрейма не должна подтверждать rename.
                _restoreRenameFocusAfterPointerActivation = false;
                this.Dispatcher.BeginInvoke(new Action(() => {
                    if (this.IsRenaming) {
                        this.RenameTextBox.Focus();
                        Keyboard.Focus(this.RenameTextBox);
                    }
                }), DispatcherPriority.ContextIdle);
                return;
            }

            this.CompleteRename();
        }

        private void CompleteRename() {
            if (!this.IsRenaming || _isCompletingRename) {
                return;
            }

            _isCompletingRename = true;
            try {
                var owner = Helpers.VisualTree.FindParentByType<TabsManagerToolWindowControl>(this);
                owner?.TryRenameTabItem(this, this.RenameText, _renameGroupTabItems);

                // Диалог ошибки сам временно забирает фокус и повторно вызывает LostKeyboardFocus.
                // Всегда завершаем F2-сценарий после одной попытки, чтобы не зациклить проверку.
                this.IsRenaming = false;
                _renameGroupTabItems = Array.Empty<State.Document.TabItemDocument>();
                _restoreRenameFocusAfterPointerActivation = false;
                this.RestoreOwnerInputTarget();
            }
            finally {
                _isCompletingRename = false;
            }
        }

        private void RestoreOwnerInputTarget() {
            Helpers.VisualTree.FindParentByType<TabsManagerToolWindowControl>(this)?.RestoreTabNavigationInputTarget();
        }

        internal void SetTabHeight(double height) {
            // Локальное значение реального контейнера немедленно инвалидирует layout.
            // Подмена ключа поверх уже разрешённого DynamicResource срабатывает только
            // после следующего пересчёта стиля (например, при смене hover/focus).
            this.TabContainer.Height = height;
        }

        private void OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e) {
            // Контекстное меню живёт в Popup, но его MouseLeftButtonUp всё ещё может дойти
            // до TabItemControl через logical tree. Меню не должно активировать вкладку.
            if (e.OriginalSource is MenuItem ||
                e.OriginalSource is DependencyObject menuSource &&
                Helpers.VisualTree.FindParentByType<MenuItem>(menuSource) != null) {

                return;
            }

            // Клик внутри поля нужен только для установки каретки/выделения. Не запускаем
            // навигацию вкладки, иначе она забирает фокус и преждевременно завершает rename.
            if (this.IsRenaming &&
                e.OriginalSource is DependencyObject renameSource &&
                (ReferenceEquals(renameSource, this.RenameTextBox) ||
                 Helpers.VisualTree.FindParentByType<TextBox>(renameSource) != null)) {
                return;
            }

            // Кнопки pin/close/keep-open находятся внутри вкладки, но их клик не является
            // выбором самой вкладки и не должен менять активный VS-фрейм.
            if (e.OriginalSource is Button ||
                e.OriginalSource is DependencyObject d && Helpers.VisualTree.FindParentByType<Button>(d) != null
                ) {
                return;
            }

            _lastPointerNavigationUtc = DateTime.UtcNow;
            Helpers.VisualTree.FindParentByType<TabsManagerToolWindowControl>(this)?.HandleTabPointerNavigation(this, Keyboard.Modifiers);
        }

        private void OnMouseEnter(object sender, MouseEventArgs e) {
            this.IsMouseInside = true;
        }

        private void OnMouseLeave(object sender, MouseEventArgs e) {
            this.IsMouseInside = false;
        }

        private void OnMouseRightButtonUpHandler(object sender, MouseButtonEventArgs e) {
            Helpers.Diagnostic.Logger.LogDebug($"  TabItemControl.OnMouseRightButtonUpHandler()");
            MenuControl? menuControl = null;

            if (_cachedWeakMenuControl?.TryGetTarget(out var cachedMenuControl) == true) {
                menuControl = cachedMenuControl;
            }
            else {
                menuControl = this.FindAndCacheMenuControl();
            }

            if (menuControl != null) {
                var mouseScreenPoint = this.ex_ToDpiAwareScreen(e.GetPosition(this));
                menuControl.ShowMenu(this.DataContext, PlacementMode.Absolute, isStaysOpen: false, mouseScreenPoint);
            }

            e.Handled = true;
        }

        private MenuControl FindAndCacheMenuControl() {
            var menuControl = Helpers.VisualTree.FindChildByType<MenuControl>(this.ContextMenuContentPresenter);
            if (menuControl == null) {
                throw new InvalidOperationException("ContextMenuTemplate must produce a MenuControl.");
            }
            _cachedWeakMenuControl = new WeakReference<MenuControl>(menuControl);
            return menuControl;
        }
    }
}
