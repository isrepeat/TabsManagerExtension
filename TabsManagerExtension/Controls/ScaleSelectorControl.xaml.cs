using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Windows.Threading;
using TabsManagerExtension.VsShell.TextEditor;

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


        public event EventHandler<double> ScaleChanged;

        private TextBox? _comboBoxTextBox;
        private double _minScale = 0.5;
        private double _maxScale = 1.5;

        public ScaleSelectorControl() {
            this.InitializeComponent();
            this.Loaded += this.OnLoaded;
            this.Unloaded += this.OnUnloaded;

            // WARNING:
            // Не присваивай this.DataContext = this — это может вызвать StackOverflow из-за биндингов вроде {Binding Text}.
            // Такие биндинги могут замкнуться на унаследованные свойства Control.Text, Content и т.п.
            // Используй ElementName / RelativeSource или выноси данные в отдельную ViewModel.
        }

        private void OnLoaded(object sender, RoutedEventArgs e) {
            Services.ExtensionServices.BeginUsage();

            this.ScaleComboBox.LostFocus += this.ScaleComboBox_OnLostFocus;
            this.ScaleComboBox.SelectionChanged += this.ScaleComboBox_OnSelectionChanged;
            VsShell.TextEditor.Services.TextEditorInputCommandFilterService.Instance.AddTrackedInputElement(this);

            // Получаем ссылку на текстовое поле внутри ComboBox (editable part)
            _comboBoxTextBox = this.ScaleComboBox.Template.FindName("PART_EditableTextBox", this.ScaleComboBox) as TextBox;
            if (_comboBoxTextBox != null) {
                _comboBoxTextBox.PreviewKeyDown += this.ScaleTextBox_OnPreviewKeyDown;
            }

            this.SelectPresetForCurrentScale();
            this.UpdateComboBoxText();
        }

        private void OnUnloaded(object sender, RoutedEventArgs e) {
            VsShell.TextEditor.Services.TextEditorInputCommandFilterService.Instance.RemoveTrackedInputElement(this);
            if (_comboBoxTextBox != null) {
                _comboBoxTextBox.PreviewKeyDown -= this.ScaleTextBox_OnPreviewKeyDown;
                _comboBoxTextBox = null;
            }

            this.ScaleComboBox.SelectionChanged -= this.ScaleComboBox_OnSelectionChanged;
            this.ScaleComboBox.LostFocus -= this.ScaleComboBox_OnLostFocus;

            Services.ExtensionServices.EndUsage();
        }


        private void ScaleComboBox_OnLostFocus(object sender, RoutedEventArgs e) {
            this.ApplyScaleFromText();
        }

        private void ScaleTextBox_OnPreviewKeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.Enter) {
                this.ApplyScaleFromText();
                e.Handled = true;
            }
        }

        private void ScaleComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (this.ScaleComboBox.SelectedItem is ComboBoxItem selectedItem && selectedItem.Tag != null) {
                if (double.TryParse(selectedItem.Tag.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double scaleFactor)) {
                    scaleFactor = Helpers.Math.Clamp(scaleFactor, _minScale, _maxScale);
                    this.CommitScale(scaleFactor);
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
            if (Math.Abs(this.ScaleFactor - scaleFactor) <= 0.001) {
                this.UpdateComboBoxText();
                return;
            }

            // SetCurrentValue сохраняет TwoWay binding с родительским контролом.
            this.SetCurrentValue(ScaleFactorProperty, scaleFactor);
            this.GetBindingExpression(ScaleFactorProperty)?.UpdateSource();
            this.ScaleChanged?.Invoke(this, this.ScaleFactor);
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

        private void SelectPresetForCurrentScale() {
            this.ScaleComboBox.SelectedItem = this.ScaleComboBox.Items
                .OfType<ComboBoxItem>()
                .FirstOrDefault(item =>
                    double.TryParse(
                        item.Tag?.ToString(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out double presetScale) &&
                    Math.Abs(presetScale - this.ScaleFactor) <= 0.001);
        }
    }
}
