using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Controls;
using System.Globalization;
using Forms = System.Windows.Forms;

namespace TabsManagerExtension.Controls {
    public partial class TabsManagerSettingsControl : UserControl {
        // Защищает обработчики контролов от обратной записи во время программного заполнения формы.
        private bool _isLoading;

        public TabsManagerSettingsControl() {
            this.InitializeComponent();
            this.ShowSection(Settings.TabsManagerSettingsService.ActiveSettingsSection, persist: false);
            this.IsEnabled = Settings.TabsManagerSettingsService.IsSettingsInitialized;
            this.Loaded += this.OnLoaded;
            this.Unloaded += this.OnUnloaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e) {
            // Страница может оставаться открытой, поэтому отражаем изменения, пришедшие из других частей расширения.
            Settings.TabsManagerSettingsService.SettingsInitialized += this.OnSettingsInitialized;
            Settings.TabsManagerSettingsService.AppearanceChanged += this.OnAppearanceChanged;
            Settings.TabsManagerSettingsService.AnchorPatternsChanged += this.OnAnchorPatternsChanged;
            Settings.TabsManagerSettingsService.AutoLoadCustomTabsChanged += this.OnAutoLoadCustomTabsChanged;
            this.IsEnabled = Settings.TabsManagerSettingsService.IsSettingsInitialized;
            this.LoadValues();
        }

        private void OnUnloaded(object sender, RoutedEventArgs e) {
            Settings.TabsManagerSettingsService.SettingsInitialized -= this.OnSettingsInitialized;
            Settings.TabsManagerSettingsService.AppearanceChanged -= this.OnAppearanceChanged;
            Settings.TabsManagerSettingsService.AnchorPatternsChanged -= this.OnAnchorPatternsChanged;
            Settings.TabsManagerSettingsService.AutoLoadCustomTabsChanged -= this.OnAutoLoadCustomTabsChanged;
        }

        private void OnSettingsInitialized() {
            this.Dispatcher.InvokeAsync(() => {
                this.IsEnabled = true;
                this.LoadValues();
            });
        }

        private void OnAppearanceChanged() {
            this.Dispatcher.InvokeAsync(this.LoadValues);
        }

        private void OnAnchorPatternsChanged() {
            this.Dispatcher.InvokeAsync(this.LoadMainValues);
        }

        private void OnAutoLoadCustomTabsChanged(bool enabled) {
            this.Dispatcher.InvokeAsync(this.LoadMainValues);
        }

        private void LoadValues() {
            // Все значения читаются одним проходом, чтобы пользователь не видел частично обновлённую секцию.
            this._isLoading = true;
            try {
                this.SetColorValue(this.PanelBackgroundText, "panelBackgroundColor", this.PanelBackgroundSwatch);
                this.SetColorValue(this.TabBackgroundText, "tabBackgroundColor", this.TabBackgroundSwatch);
                this.SetColorValue(this.TabBorderText, "tabBorderColor", this.TabBorderSwatch);
                this.SetColorValue(this.TabHoverBackgroundText, "tabHoverBackgroundColor", this.TabHoverBackgroundSwatch);
                this.SetColorValue(this.TabHoverBorderText, "tabHoverBorderColor", this.TabHoverBorderSwatch);
                this.SetColorValue(this.SelectedTabBackgroundText, "selectedTabBackgroundColor", this.SelectedTabBackgroundSwatch);
                this.SetColorValue(this.SelectedTabBorderText, "selectedTabBorderColor", this.SelectedTabBorderSwatch);
                this.SetColorValue(this.ActiveTabBackgroundText, "activeTabBackgroundColor", this.ActiveTabBackgroundSwatch);
                this.SetColorValue(this.ActiveTabBorderText, "activeTabBorderColor", this.ActiveTabBorderSwatch);
                this.SetColorValue(this.TabTextColorText, "tabTextColor", this.TabTextColorSwatch);
                this.SetColorValue(this.TabHoverTextColorText, "tabHoverTextColor", this.TabHoverTextColorSwatch);
                this.SetColorValue(this.SelectedTabTextColorText, "selectedTabTextColor", this.SelectedTabTextColorSwatch);
                this.SetColorValue(this.ActiveTabTextColorText, "activeTabTextColor", this.ActiveTabTextColorSwatch);
                this.SetFontValues(this.TabWeight, this.TabStyle, this.TabSize, "tabTextBold", "tabTextItalic", "tabTextSize");
                this.SetFontValues(this.TabHoverWeight, this.TabHoverStyle, this.TabHoverSize, "tabHoverTextBold", "tabHoverTextItalic", "tabHoverTextSize");
                this.SetFontValues(this.SelectedTabWeight, this.SelectedTabStyle, this.SelectedTabSize, "selectedTabTextBold", "selectedTabTextItalic", "selectedTabTextSize");
                this.SetFontValues(this.ActiveTabWeight, this.ActiveTabStyle, this.ActiveTabSize, "activeTabTextBold", "activeTabTextItalic", "activeTabTextSize");
                this.LoadMainValuesCore();
            }
            finally {
                this._isLoading = false;
            }
        }

        private void LoadMainValues() {
            this._isLoading = true;
            try {
                this.LoadMainValuesCore();
            }
            finally {
                this._isLoading = false;
            }
        }

        private void LoadMainValuesCore() {
            this.AutoLoadTabsCheckBox.IsChecked = Settings.TabsManagerSettingsService.AutoLoadCustomTabs;
            this.LoggingEnabledCheckBox.IsChecked = Settings.TabsManagerSettingsService.IsLoggingEnabled;
            this.LoggingSessionModeComboBox.SelectedValue = Settings.TabsManagerSettingsService.LoggingSessionMode;
            this.UpdateLoggingSessionModeRestartMessage();
            this.AnchorSectionPatternText.Text = Settings.TabsManagerSettingsService.AnchorSectionPattern;
            this.AnchorSubsectionPatternText.Text = Settings.TabsManagerSettingsService.AnchorSubsectionPattern;
        }

        private void OnMainNavigationClicked(object sender, RoutedEventArgs e) {
            this.ShowSection("main", persist: true);
        }

        private void OnCustomizationNavigationClicked(object sender, RoutedEventArgs e) {
            this.ShowSection("customization", persist: true);
        }

        private void OnAnchorsNavigationClicked(object sender, RoutedEventArgs e) {
            this.ShowSection("anchors", persist: true);
        }

        private void ShowSection(string section, bool persist) {
            // Выбранная секция сохраняется в settings.json и восстанавливается при следующем открытии вкладки.
            bool showCustomization = string.Equals(section, "customization", StringComparison.Ordinal);
            bool showAnchors = string.Equals(section, "anchors", StringComparison.Ordinal);
            string normalizedSection = showCustomization ? "customization" : showAnchors ? "anchors" : "main";

            this.MainPanel.Visibility = normalizedSection == "main" ? Visibility.Visible : Visibility.Collapsed;
            this.CustomizationPanel.Visibility = showCustomization ? Visibility.Visible : Visibility.Collapsed;
            this.AnchorsPanel.Visibility = showAnchors ? Visibility.Visible : Visibility.Collapsed;
            this.MainNavigationButton.Tag = normalizedSection == "main" ? "Selected" : null;
            this.CustomizationNavigationButton.Tag = showCustomization ? "Selected" : null;
            this.AnchorsNavigationButton.Tag = showAnchors ? "Selected" : null;

            if (persist) {
                Settings.TabsManagerSettingsService.SetActiveSettingsSection(normalizedSection);
            }
        }

        private void OnAutoLoadTabsChanged(object sender, RoutedEventArgs e) {
            if (!this._isLoading) {
                Settings.TabsManagerSettingsService.SetAutoLoadCustomTabs(this.AutoLoadTabsCheckBox.IsChecked == true);
            }
        }

        private void OnLoggingEnabledChanged(object sender, RoutedEventArgs e) {
            if (!this._isLoading) {
                Settings.TabsManagerSettingsService.SetLoggingEnabled(this.LoggingEnabledCheckBox.IsChecked == true);
            }
        }

        private void OnLoggingSessionModeChanged(object sender, SelectionChangedEventArgs e) {
            if (this._isLoading || this.LoggingSessionModeComboBox.SelectedValue is not string loggingSessionMode) {
                return;
            }

            Settings.TabsManagerSettingsService.SetLoggingSessionMode(loggingSessionMode);
            this.UpdateLoggingSessionModeRestartMessage();
        }

        private void UpdateLoggingSessionModeRestartMessage() {
            bool modeChanged = this.LoggingSessionModeComboBox.SelectedValue is string loggingSessionMode &&
                !string.Equals(
                    loggingSessionMode,
                    Settings.TabsManagerSettingsService.CurrentLoggingSessionMode,
                    StringComparison.Ordinal
                );
            this.LoggingSessionModeRestartMessage.Visibility = modeChanged ? Visibility.Visible : Visibility.Collapsed;
        }

        private void OnAnchorPatternCommitted(object sender, KeyboardFocusChangedEventArgs e) {
            if (sender is TextBox textBox) {
                this.CommitAnchorPattern(textBox);
            }
        }

        private void CommitAnchorPattern(TextBox textBox) {
            // Сервис валидирует регулярное выражение и при ошибке возвращает безопасный шаблон по умолчанию.
            if (string.Equals(textBox.Tag as string, "section", StringComparison.Ordinal)) {
                Settings.TabsManagerSettingsService.SetAnchorSectionPattern(textBox.Text);
            }
            else if (string.Equals(textBox.Tag as string, "subsection", StringComparison.Ordinal)) {
                Settings.TabsManagerSettingsService.SetAnchorSubsectionPattern(textBox.Text);
            }

            this.LoadMainValues();
        }

        private void SetColorValue(TextBox textBox, string key, Border? swatch = null) {
            string value = Settings.TabsManagerSettingsService.GetAppearanceColor(key);
            textBox.Text = value;
            if (key.IndexOf("TextColor", StringComparison.Ordinal) < 0) {
                textBox.Width = 150;
                textBox.HorizontalAlignment = HorizontalAlignment.Left;
            }

            if (swatch != null) {
                swatch.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(value));
            }
        }

        private void SetFontValues(CheckBox weight, CheckBox style, TextBox size, string weightKey, string styleKey, string sizeKey) {
            weight.IsChecked = Settings.TabsManagerSettingsService.GetAppearanceBoolean(weightKey);
            style.IsChecked = Settings.TabsManagerSettingsService.GetAppearanceBoolean(styleKey);
            size.Text = Settings.TabsManagerSettingsService.GetAppearanceNumber(sizeKey).ToString(CultureInfo.CurrentCulture);
        }

        private void OnColorTextCommitted(object sender, KeyboardFocusChangedEventArgs e) {
            if (sender is TextBox textBox && textBox.Tag is string key) {
                Settings.TabsManagerSettingsService.SetAppearanceColor(key, textBox.Text);
                this.LoadValues();
            }
        }

        private void OnColorSwatchClicked(object sender, MouseButtonEventArgs e) {
            if (sender is not Border swatch || swatch.Tag is not string textBoxName || this.FindName(textBoxName) is not TextBox textBox) {
                return;
            }

            var currentColor = (Color)ColorConverter.ConvertFromString(textBox.Text);
            using var dialog = new Forms.ColorDialog {
                FullOpen = true,
                Color = System.Drawing.Color.FromArgb(currentColor.A, currentColor.R, currentColor.G, currentColor.B)
            };

            if (dialog.ShowDialog() != Forms.DialogResult.OK) {
                return;
            }

            textBox.Text = $"#{dialog.Color.A:X2}{dialog.Color.R:X2}{dialog.Color.G:X2}{dialog.Color.B:X2}";
            Settings.TabsManagerSettingsService.SetAppearanceColor((string)textBox.Tag, textBox.Text);
            this.LoadValues();
        }

        private void OnResetSettingsClicked(object sender, RoutedEventArgs e) {
            Settings.TabsManagerSettingsService.ResetAppearanceToDefaults();
            this.LoadValues();
        }

        private void OnFontSettingChanged(object sender, RoutedEventArgs e) {
            if (this._isLoading || sender is not CheckBox checkBox || checkBox.Tag is not string key) {
                return;
            }

            Settings.TabsManagerSettingsService.SetAppearanceBoolean(key, checkBox.IsChecked == true);
        }

        private void OnEditorKeyDown(object sender, KeyEventArgs e) {
            if (e.Key != Key.Enter || sender is not TextBox textBox) {
                return;
            }

            if (textBox.Tag is string key && key.EndsWith("Color", StringComparison.Ordinal)) {
                Settings.TabsManagerSettingsService.SetAppearanceColor(key, textBox.Text);
                this.LoadValues();
            }
            else if (string.Equals(textBox.Tag as string, "section", StringComparison.Ordinal) ||
                     string.Equals(textBox.Tag as string, "subsection", StringComparison.Ordinal)) {
                this.CommitAnchorPattern(textBox);
            }
            else {
                this.CommitFontSize(textBox);
            }

            e.Handled = true;
        }

        private void OnFontSizeCommitted(object sender, KeyboardFocusChangedEventArgs e) {
            if (sender is TextBox textBox) {
                this.CommitFontSize(textBox);
            }

            this.LoadValues();
        }

        private void CommitFontSize(TextBox textBox) {
            if (textBox.Tag is string key &&
                double.TryParse(textBox.Text, NumberStyles.Float, CultureInfo.CurrentCulture, out double value)) {
                Settings.TabsManagerSettingsService.SetAppearanceNumber(key, value);
            }
        }
    }
}
