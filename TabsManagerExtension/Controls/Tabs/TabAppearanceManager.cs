using System.Windows;
using System.Windows.Media;
using System.Collections.Generic;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Применяет масштаб и ресурсы appearance к корневому control и его вкладкам.</summary>
    internal sealed class TabAppearanceManager {
        public void ApplyScale(TabItemControl control, double scaleFactor) {
            // База 26.4 соответствует прежней высоте вкладки при масштабе 120%.
            control.SetTabHeight(scaleFactor * 26.4);
        }

        public void ApplyAppearance(
            ResourceDictionary rootResources,
            IEnumerable<TabItemControl> tabItemControls
            ) {
            this.ApplyResources(rootResources);

            foreach (var tabItemControl in tabItemControls) {
                this.ApplyResources(tabItemControl.Resources);
            }
        }

        public void ApplyAppearance(TabItemControl control) {
            this.ApplyResources(control.Resources);
        }

        private void ApplyResources(ResourceDictionary resources) {
            // Одни и те же динамические ресурсы записываются в корневой контрол и в уже созданные
            // TabItemControl, поэтому изменения настроек видны без пересоздания списка.
            resources["AppTabsPanelBackgroundBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("panelBackgroundColor"));
            resources["AppTabBackgroundBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("tabBackgroundColor"));
            resources["AppTabBorderBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("tabBorderColor"));
            resources["AppTabHoverBackgroundBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("tabHoverBackgroundColor"));
            resources["AppTabHoverBorderBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("tabHoverBorderColor"));
            resources["AppTabSelectedBackgroundBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("selectedTabBackgroundColor"));
            resources["AppTabSelectedBorderBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("selectedTabBorderColor"));
            resources["AppTabActiveBackgroundBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("activeTabBackgroundColor"));
            resources["AppTabActiveBorderBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("activeTabBorderColor"));
            resources["AppTabForegroundBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("tabTextColor"));
            resources["AppTabHoverForegroundBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("tabHoverTextColor"));
            resources["AppTabSelectedForegroundBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("selectedTabTextColor"));
            resources["AppTabActiveForegroundBrush"] = CreateBrush(Settings.TabsManagerSettingsService.GetAppearanceColor("activeTabTextColor"));
            resources["AppTabFontWeight"] = Settings.TabsManagerSettingsService.GetAppearanceBoolean("tabTextBold") ? FontWeights.Bold : FontWeights.Normal;
            resources["AppTabFontStyle"] = Settings.TabsManagerSettingsService.GetAppearanceBoolean("tabTextItalic") ? FontStyles.Italic : FontStyles.Normal;
            resources["AppTabItemFontSize"] = Settings.TabsManagerSettingsService.GetAppearanceNumber("tabTextSize");
            resources["AppTabHoverFontWeight"] = Settings.TabsManagerSettingsService.GetAppearanceBoolean("tabHoverTextBold") ? FontWeights.Bold : FontWeights.Normal;
            resources["AppTabHoverFontStyle"] = Settings.TabsManagerSettingsService.GetAppearanceBoolean("tabHoverTextItalic") ? FontStyles.Italic : FontStyles.Normal;
            resources["AppTabHoverFontSize"] = Settings.TabsManagerSettingsService.GetAppearanceNumber("tabHoverTextSize");
            resources["AppTabSelectedFontWeight"] = Settings.TabsManagerSettingsService.GetAppearanceBoolean("selectedTabTextBold") ? FontWeights.Bold : FontWeights.Normal;
            resources["AppTabSelectedFontStyle"] = Settings.TabsManagerSettingsService.GetAppearanceBoolean("selectedTabTextItalic") ? FontStyles.Italic : FontStyles.Normal;
            resources["AppTabSelectedFontSize"] = Settings.TabsManagerSettingsService.GetAppearanceNumber("selectedTabTextSize");
            resources["AppTabActiveFontWeight"] = Settings.TabsManagerSettingsService.GetAppearanceBoolean("activeTabTextBold") ? FontWeights.Bold : FontWeights.Normal;
            resources["AppTabActiveFontStyle"] = Settings.TabsManagerSettingsService.GetAppearanceBoolean("activeTabTextItalic") ? FontStyles.Italic : FontStyles.Normal;
            resources["AppTabActiveFontSize"] = Settings.TabsManagerSettingsService.GetAppearanceNumber("activeTabTextSize");
        }

        private static SolidColorBrush CreateBrush(string value) {
            return new SolidColorBrush((Color)ColorConverter.ConvertFromString(value));
        }
    }
}
