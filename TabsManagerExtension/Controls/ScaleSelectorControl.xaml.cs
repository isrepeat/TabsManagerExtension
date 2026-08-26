using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;

namespace TabsManagerExtension.Controls {
    public partial class ScaleSelectorControl : Helpers.BaseUserControl {
        public string Title {
            get { return (string)this.GetValue(TitleProperty); }
            set { this.SetValue(TitleProperty, value); }
        }
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(
                nameof(Title),
                typeof(string),
                typeof(ScaleSelectorControl),
                new PropertyMetadata(""));

        public double ScaleFactor {
            get => (double)GetValue(ScaleFactorProperty);
            set => SetValue(ScaleFactorProperty, value);
        }
        public static readonly DependencyProperty ScaleFactorProperty =
            DependencyProperty.Register(
                nameof(ScaleFactor),
                typeof(double),
                typeof(ScaleSelectorControl),
                new PropertyMetadata(1.0));


        private TextBox? _comboBoxTextBox;
        private double _minScale = 0.5;
        private double _maxScale = 1.5;

        public ScaleSelectorControl() {
            this.InitializeComponent();
            this.Loaded += this.OnLoaded;
            this.Unloaded += this.OnUnloaded;

            // Не назначаем DataContext = this: контрол должен наследовать DataContext родителя,
            // чтобы внешние привязки, например ScaleFactorTabsCompactness, продолжали работать.
            // Для обращения к свойствам самого контрола в XAML используем ElementName или RelativeSource.
        }

        private void OnLoaded(object sender, RoutedEventArgs e) {
            Services.ExtensionServices.BeginUsage();

            this.ScaleComboBox.LostFocus += this.ScaleComboBox_OnLostFocus;
            // Поле ввода внутри ComboBox создаётся самим WPF и может быть заменено, например,
            // после смены темы. Тогда обработчик, подписанный прямо на старое поле ввода,
            // перестанет работать. Поэтому обрабатываем Enter на самом ComboBox: нажатие
            // сначала получает внутреннее поле, а затем WPF передаёт его родительскому ComboBox.
            this.ScaleComboBox.KeyDown += this.ScaleComboBox_OnKeyDown;
            this.ScaleComboBox.SelectionChanged += this.ScaleComboBox_OnSelectionChanged;
            VsShell.TextEditor.Services.TextEditorInputCommandFilterService.Instance.AddTrackedInputElement(this);

            // Получаем ссылку на текстовое поле внутри ComboBox (editable part)
            _comboBoxTextBox = this.ScaleComboBox.Template.FindName("PART_EditableTextBox", this.ScaleComboBox) as TextBox;
            this.UpdateComboBoxText();
        }

        private void OnUnloaded(object sender, RoutedEventArgs e) {
            VsShell.TextEditor.Services.TextEditorInputCommandFilterService.Instance.RemoveTrackedInputElement(this);
            this.ScaleComboBox.SelectionChanged -= this.ScaleComboBox_OnSelectionChanged;
            this.ScaleComboBox.KeyDown -= this.ScaleComboBox_OnKeyDown;
            this.ScaleComboBox.LostFocus -= this.ScaleComboBox_OnLostFocus;
            _comboBoxTextBox = null;

            Services.ExtensionServices.EndUsage();
        }


        private void ScaleComboBox_OnLostFocus(object sender, RoutedEventArgs e) {
            this.ApplyScaleFromText();
        }

        private void ScaleComboBox_OnKeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.Enter) {
                this.ApplyScaleFromText();
                e.Handled = true;
            }
        }

        private void ScaleComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (this.ScaleComboBox.SelectedItem is ComboBoxItem selectedItem && selectedItem.Tag != null) {
                if (double.TryParse(selectedItem.Tag.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double scaleFactor)) {
                    this.CommitScale(Helpers.Math.Clamp(scaleFactor, _minScale, _maxScale));
                }
            }
        }

        private void ApplyScaleFromText() {
            if (_comboBoxTextBox != null) {
                string input = _comboBoxTextBox.Text.Replace("%", "").Trim();

                if (double.TryParse(input, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double scaleValue)) {
                    scaleValue = Helpers.Math.Clamp(scaleValue / 100.0, _minScale, _maxScale);
                    this.CommitScale(scaleValue);
                }
                else {
                    this.UpdateComboBoxText();
                }

                this.ScaleComboBox.SelectedItem = null;
            }
        }

        private void CommitScale(double scaleFactor) {
            if (Math.Abs(this.ScaleFactor - scaleFactor) > 0.001) {
                // SetCurrentValue сохраняет TwoWay binding с родительским контролом.
                this.SetCurrentValue(ScaleFactorProperty, scaleFactor);
                this.GetBindingExpression(ScaleFactorProperty)?.UpdateSource();
            }

            this.UpdateComboBoxText();
        }

        private void UpdateComboBoxText() {
            if (_comboBoxTextBox != null) {
                string newText = (this.ScaleFactor * 100).ToString("F0") + " %";

                if (_comboBoxTextBox.Text != newText) {
                    _comboBoxTextBox.Text = newText;

                    // Ставим каретку перед символом '%'
                    int caretPos = newText.LastIndexOf('%');
                    if (caretPos > 0) {
                        _comboBoxTextBox.CaretIndex = caretPos - 1; // перед пробелом
                    }
                }
            }
        }

    }
}
