using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Collections.ObjectModel;
using FormsScreen = System.Windows.Forms.Screen;
using DrawingPoint = System.Drawing.Point;

namespace TabsManagerExtension.Controls {
    public partial class MenuControl : Helpers.BaseUserControl {
        public ObservableCollection<Helpers.IMenuItem> MenuItems {
            get => (ObservableCollection<Helpers.IMenuItem>)this.GetValue(MenuItemsProperty);
            set => this.SetValue(MenuItemsProperty, value);
        }
        public static readonly DependencyProperty MenuItemsProperty =
            DependencyProperty.Register(
                nameof(MenuItems),
                typeof(ObservableCollection<Helpers.IMenuItem>),
                typeof(MenuControl),
                new PropertyMetadata(null));


        public ICommand OpenCommand {
            get => (ICommand)this.GetValue(OpenCommandProperty);
            set => this.SetValue(OpenCommandProperty, value);
        }
        public static readonly DependencyProperty OpenCommandProperty =
            DependencyProperty.Register(
                nameof(OpenCommand),
                typeof(ICommand),
                typeof(MenuControl),
                new PropertyMetadata(null));


        public ICommand CloseCommand {
            get => (ICommand)this.GetValue(CloseCommandProperty);
            set => this.SetValue(CloseCommandProperty, value);
        }
        public static readonly DependencyProperty CloseCommandProperty =
            DependencyProperty.Register(
                nameof(CloseCommand),
                typeof(ICommand),
                typeof(MenuControl),
                new PropertyMetadata(null));


        public DataTemplateSelector? MenuItemTemplateSelector {
            get => (DataTemplateSelector?)this.GetValue(MenuItemTemplateSelectorProperty);
            set => this.SetValue(MenuItemTemplateSelectorProperty, value);
        }
        public static readonly DependencyProperty MenuItemTemplateSelectorProperty =
            DependencyProperty.Register(
                nameof(MenuItemTemplateSelector),
                typeof(DataTemplateSelector),
                typeof(MenuControl),
                new PropertyMetadata(null));


        public class MenuOpeningArgs {
            public object DataContext { get; set; } = default!;
            public bool ShouldOpen { get; set; } = true;
        }
        public class MenuClosedArgs {
            public object DataContext { get; set; } = default!;
        }

        public event EventHandler MouseEnteredPopup;
        public event EventHandler MouseLeftPopup;

        private MenuClosedArgs? _menuControlClosedArgs;
        // Хранит пункт, выбранный между событиями MenuItem.Click и Popup.Closed.
        // Это позволяет сначала полностью скрыть меню, а затем выполнить команду пункта;
        // при обычном закрытии без выбора значение остаётся null и команда не запускается.
        private Helpers.MenuItemCommand? _pendingMenuItemCommand;

        public MenuControl() {
            this.InitializeComponent();
            this.MenuItemsControl.AddHandler(MenuItem.ClickEvent, new RoutedEventHandler(this.MenuItem_OnClick), true);
            this.MouseLeftButtonDown += this.OnBlockMouseBubbling;
            this.MouseLeftButtonUp += this.OnBlockMouseBubbling;
            this.MouseRightButtonDown += this.OnBlockMouseBubbling;
            this.MouseRightButtonUp += this.OnBlockMouseBubbling;
        }

        public void ShowMenu(
            object dataContext,
            PlacementMode placementMode,
            bool isStaysOpen,
            Point? screenPosition = null,
            FrameworkElement? placementTarget = null
        ) {
            _menuControlClosedArgs = new MenuControl.MenuClosedArgs {
                DataContext = dataContext
            };

            if (this.OpenCommand != null && this.OpenCommand.CanExecute(null)) {
                var request = new MenuControl.MenuOpeningArgs {
                    DataContext = dataContext
                };

                // Через команду заполняются MenuItems.
                this.OpenCommand.Execute(request);
                if (!request.ShouldOpen) {
                    _menuControlClosedArgs = null;
                    return;
                }
            }

            this.MenuPopup.Placement = placementMode;
            this.MenuPopup.PlacementTarget = placementTarget;
            this.MenuPopup.StaysOpen = isStaysOpen;
            if (screenPosition.HasValue) {
                this.UpdateMaximumMenuHeight(screenPosition.Value);
                this.MoveMenu(screenPosition.Value);
            }
            else {
                this.MenuScrollViewer.MaxHeight = Math.Max(0, SystemParameters.WorkArea.Height - 20);
            }
            this.MenuPopup.IsOpen = true;
        }


        public void UpdateMenu(object dataContext, Point? screenPosition = null) {
            _menuControlClosedArgs = new MenuControl.MenuClosedArgs {
                DataContext = dataContext
            };

            if (this.OpenCommand != null && this.OpenCommand.CanExecute(null)) {
                var request = new MenuControl.MenuOpeningArgs {
                    DataContext = dataContext
                };

                // Через команду заполняются MenuItems.
                this.OpenCommand.Execute(request);

                // Do not handle request.ShouldOpen here.
            }

            if (screenPosition.HasValue) {
                this.UpdateMaximumMenuHeight(screenPosition.Value);
                this.MoveMenu(screenPosition.Value);
            }
        }

        private void UpdateMaximumMenuHeight(Point screenPosition) {
            var source = PresentationSource.FromVisual(this);
            double dpiScaleX = source?.CompositionTarget?.TransformToDevice.M11 ?? 1;
            double dpiScaleY = source?.CompositionTarget?.TransformToDevice.M22 ?? 1;
            var physicalPoint = new DrawingPoint(
                (int)Math.Round(screenPosition.X * dpiScaleX),
                (int)Math.Round(screenPosition.Y * dpiScaleY)
            );

            var workArea = FormsScreen.FromPoint(physicalPoint).WorkingArea;
            double availableAbove = (physicalPoint.Y - workArea.Top) / dpiScaleY;
            double availableBelow = (workArea.Bottom - physicalPoint.Y) / dpiScaleY;

            // Popup сам выбирает сторону размещения; ограничиваем меню большей из доступных областей.
            this.MenuScrollViewer.MaxHeight = Math.Max(0, Math.Max(availableAbove, availableBelow) - 20);
        }

        public void MoveMenu(Point screenPosition) {
            this.MenuPopup.HorizontalOffset = screenPosition.X;
            this.MenuPopup.VerticalOffset = screenPosition.Y;
        }



        private void OnBlockMouseBubbling(object sender, MouseButtonEventArgs e) {
            e.Handled = true;
        }

        private void MenuItem_OnClick(object sender, RoutedEventArgs e) {
            if (e.OriginalSource is not MenuItem { DataContext: Helpers.MenuItemCommand menuItemCommand }) {
                return;
            }

            // В шаблоне MenuItem нет прямой привязки Command, поэтому клик только инициирует
            // закрытие. Команда будет выполнена из MenuPopup_OnClosed.
            _pendingMenuItemCommand = menuItemCommand;
            this.MenuPopup.IsOpen = false;
        }

        private void MenuPopup_OnClosed(object sender, System.EventArgs e) {
            var closedArgs = _menuControlClosedArgs;
            _menuControlClosedArgs = null;
            if (closedArgs == null) {
                return;
            }

            if (this.CloseCommand != null && this.CloseCommand.CanExecute(closedArgs)) {
                this.CloseCommand.Execute(closedArgs);
            }

            // Popup уже полностью скрыт, поэтому даже тяжёлая команда больше не задержит
            // исчезновение меню на экране.
            var menuItemCommand = _pendingMenuItemCommand;
            _pendingMenuItemCommand = null;
            if (menuItemCommand?.Command.CanExecute(menuItemCommand.CommandParameterContext) == true) {
                menuItemCommand.Command.Execute(menuItemCommand.CommandParameterContext);
            }
        }
        private void MenuPopup_OnMouseEnter(object sender, MouseEventArgs e) {
            this.MouseEnteredPopup?.Invoke(this, EventArgs.Empty);
        }
        private void MenuPopup_OnMouseLeave(object sender, MouseEventArgs e) {
            this.MouseLeftPopup?.Invoke(this, EventArgs.Empty);
        }


        public class MenuItemTemplateSelectorBase : DataTemplateSelector {
            public DataTemplate? HeaderTemplate { get; set; }
            public DataTemplate? CommandTemplate { get; set; }
            public DataTemplate? SeparatorTemplate { get; set; }

            public override DataTemplate? SelectTemplate(object item, DependencyObject container) {
                if (item is Helpers.MenuItemHeader) {
                    return this.HeaderTemplate;
                }
                if (item is Helpers.MenuItemCommand) {
                    return this.CommandTemplate;
                }
                if (item is Helpers.MenuItemSeparator) {
                    return this.SeparatorTemplate;
                }

                return base.SelectTemplate(item, container);
            }
        }
    }
}
