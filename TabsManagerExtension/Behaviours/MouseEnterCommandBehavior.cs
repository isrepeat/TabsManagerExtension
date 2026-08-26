using System.Windows;
using System.Windows.Input;

namespace TabsManagerExtension.Behaviours {
    public static class MouseEnterCommandBehavior {
        public static readonly DependencyProperty CommandProperty = DependencyProperty.RegisterAttached(
            "Command",
            typeof(ICommand),
            typeof(MouseEnterCommandBehavior),
            new PropertyMetadata(null, OnCommandChanged)
        );

        public static readonly DependencyProperty CommandParameterProperty = DependencyProperty.RegisterAttached(
            "CommandParameter",
            typeof(object),
            typeof(MouseEnterCommandBehavior),
            new PropertyMetadata(null)
        );

        public static void SetCommand(DependencyObject element, ICommand value) {
            element.SetValue(CommandProperty, value);
        }

        public static ICommand GetCommand(DependencyObject element) {
            return (ICommand)element.GetValue(CommandProperty);
        }

        public static void SetCommandParameter(DependencyObject element, object value) {
            element.SetValue(CommandParameterProperty, value);
        }

        public static object GetCommandParameter(DependencyObject element) {
            return element.GetValue(CommandParameterProperty);
        }

        private static void OnCommandChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e) {
            if (dependencyObject is not UIElement element) {
                return;
            }

            if (e.OldValue != null) {
                element.MouseEnter -= OnMouseEnter;
            }

            if (e.NewValue != null) {
                element.MouseEnter += OnMouseEnter;
            }
        }

        private static void OnMouseEnter(object sender, MouseEventArgs e) {
            if (sender is not DependencyObject element) {
                return;
            }

            var command = GetCommand(element);
            var parameter = GetCommandParameter(element);
            if (command?.CanExecute(parameter) == true) {
                command.Execute(parameter);
            }
        }
    }
}
