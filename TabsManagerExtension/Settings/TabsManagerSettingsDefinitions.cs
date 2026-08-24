#pragma warning disable VSEXTPREVIEW_SETTINGS

using Microsoft.VisualStudio.Extensibility;
using Microsoft.VisualStudio.Extensibility.Settings;

namespace TabsManagerExtension.Settings {
    internal static class TabsManagerSettingsDefinitions {
        [VisualStudioContribution]
        internal static SettingCategory TabsManagerCategory { get; } = new SettingCategory(
            "tabsManager",
            "Tabs Manager"
        ) {
            Description = "Open the dedicated Tabs Manager settings tab to configure the extension."
        };

        private static SettingRule InternalSettingsVisibleRule { get; } =
            SettingRule.EnvironmentVariableEqual("TABS_MANAGER_SHOW_INTERNAL_SETTINGS", "1");

        [VisualStudioContribution]
        internal static Setting.Boolean OpenSettingsPage { get; } = new Setting.Boolean(
            "openSettingsPage",
            "Open Tabs Manager settings…",
            TabsManagerCategory,
            defaultValue: false
        ) {
            Description = "Opens the complete Tabs Manager settings interface in a dedicated tab."
        };

        [VisualStudioContribution]
        internal static SettingCategory AnchorsCategory { get; } = new SettingCategory(
            "anchors",
            "Comment anchors",
            TabsManagerCategory
        ) {
            Description = "Recognition patterns for anchors displayed over the text editor.",
            VisibleWhen = InternalSettingsVisibleRule
        };

        [VisualStudioContribution]
        internal static SettingCategory CustomizationCategory { get; } = new SettingCategory(
            "customization",
            "Customization",
            TabsManagerCategory
        ) {
            Description = "Colors and typography of the tabs panel. Colors use #AARRGGBB or #RRGGBB format.",
            VisibleWhen = InternalSettingsVisibleRule
        };

        [VisualStudioContribution]
        internal static Setting.Boolean AutoLoadCustomTabs { get; } = new Setting.Boolean(
            "autoLoadCustomTabs",
            "Use Tabs Manager tabs",
            TabsManagerCategory,
            defaultValue: true
        ) {
            Description = "Switches between the standard Visual Studio tabs and Tabs Manager tabs."
        };

        [VisualStudioContribution]
        internal static Setting.Boolean ShowTabsToggleToolbarButton { get; } = new Setting.Boolean(
            "showTabsToggleToolbarButton",
            "Show the Tabs Manager switch on the standard toolbar",
            TabsManagerCategory,
            defaultValue: true
        );

        [VisualStudioContribution]
        internal static Setting.Boolean ShowStandardTabsLayoutToolbarButton { get; } = new Setting.Boolean(
            "showStandardTabsLayoutToolbarButton",
            "Show the standard tabs position switch on the standard toolbar",
            TabsManagerCategory,
            defaultValue: true
        );

        [VisualStudioContribution]
        internal static Setting.Decimal TabsScaleFactor { get; } = new Setting.Decimal(
            "tabsScaleFactor",
            "Tab compactness",
            TabsManagerCategory,
            defaultValue: 1.0m
        ) {
            Description = "Controls the size and spacing of tabs.",
            Minimum = 0.5m,
            Maximum = 1.5m,
            VisibleWhen = InternalSettingsVisibleRule
        };

        [VisualStudioContribution]
        internal static Setting.String AnchorSectionPattern { get; } = new Setting.String(
            "anchorSectionPattern",
            "Anchor section pattern",
            AnchorsCategory,
            TabsManagerSettingsService.DefaultAnchorSectionPattern
        ) {
            Description = "Regular expression for a top-level anchor. Put its displayed text in the named group 'title'."
        };

        [VisualStudioContribution]
        internal static Setting.String AnchorSubsectionPattern { get; } = new Setting.String(
            "anchorSubsectionPattern",
            "Anchor subsection pattern",
            AnchorsCategory,
            TabsManagerSettingsService.DefaultAnchorSubsectionPattern
        ) {
            Description = "Regular expression for a nested anchor. Put its displayed text in the named group 'title'."
        };

        [VisualStudioContribution]
        internal static Setting.String PanelBackgroundColor { get; } = new Setting.String(
            "panelBackgroundColor",
            "Background",
            CustomizationCategory,
            "#FF252526"
        );

        [VisualStudioContribution]
        internal static Setting.String TabBackgroundColor { get; } = new Setting.String(
            "tabBackgroundColor",
            "Normal — background",
            CustomizationCategory,
            "#00FFFFFF"
        );

        [VisualStudioContribution]
        internal static Setting.String TabBorderColor { get; } = new Setting.String(
            "tabBorderColor",
            "Normal — border",
            CustomizationCategory,
            "#00FFFFFF"
        );

        [VisualStudioContribution]
        internal static Setting.String TabHoverBackgroundColor { get; } = new Setting.String(
            "tabHoverBackgroundColor",
            "Hovered — background",
            CustomizationCategory,
            "#663C3C3C"
        );

        [VisualStudioContribution]
        internal static Setting.String TabHoverBorderColor { get; } = new Setting.String(
            "tabHoverBorderColor",
            "Hovered — border",
            CustomizationCategory,
            "#00FFFFFF"
        );

        [VisualStudioContribution]
        internal static Setting.String SelectedTabBackgroundColor { get; } = new Setting.String(
            "selectedTabBackgroundColor",
            "Selected — background",
            CustomizationCategory,
            "#663C3C3C"
        );

        [VisualStudioContribution]
        internal static Setting.String SelectedTabBorderColor { get; } = new Setting.String(
            "selectedTabBorderColor",
            "Selected — border",
            CustomizationCategory,
            "#00FFFFFF"
        );

        [VisualStudioContribution]
        internal static Setting.String ActiveTabBackgroundColor { get; } = new Setting.String(
            "activeTabBackgroundColor",
            "Active — background",
            CustomizationCategory,
            "#FF2D2D30"
        );

        [VisualStudioContribution]
        internal static Setting.String ActiveTabBorderColor { get; } = new Setting.String(
            "activeTabBorderColor",
            "Active — border",
            CustomizationCategory,
            "#FF4B4B4B"
        );

        [VisualStudioContribution]
        internal static Setting.String TabTextColor { get; } = new Setting.String(
            "tabTextColor",
            "Normal — color",
            CustomizationCategory,
            "#FF808080"
        );

        [VisualStudioContribution]
        internal static Setting.Boolean TabTextBold { get; } = new Setting.Boolean(
            "tabTextBold",
            "Normal — bold",
            CustomizationCategory,
            false
        );

        [VisualStudioContribution]
        internal static Setting.Boolean TabTextItalic { get; } = new Setting.Boolean(
            "tabTextItalic",
            "Normal — italic",
            CustomizationCategory,
            false
        );

        [VisualStudioContribution]
        internal static Setting.Decimal TabTextSize { get; } = new Setting.Decimal(
            "tabTextSize",
            "Normal — size",
            CustomizationCategory,
            12m
        ) {
            Minimum = 8m,
            Maximum = 32m
        };

        [VisualStudioContribution]
        internal static Setting.String TabHoverTextColor { get; } = new Setting.String(
            "tabHoverTextColor",
            "Hovered — color",
            CustomizationCategory,
            "#FFFFFFFF"
        );

        [VisualStudioContribution]
        internal static Setting.Boolean TabHoverTextBold { get; } = new Setting.Boolean(
            "tabHoverTextBold",
            "Hovered — bold",
            CustomizationCategory,
            false
        );

        [VisualStudioContribution]
        internal static Setting.Boolean TabHoverTextItalic { get; } = new Setting.Boolean(
            "tabHoverTextItalic",
            "Hovered — italic",
            CustomizationCategory,
            false
        );

        [VisualStudioContribution]
        internal static Setting.Decimal TabHoverTextSize { get; } = new Setting.Decimal(
            "tabHoverTextSize",
            "Hovered — size",
            CustomizationCategory,
            12m
        ) {
            Minimum = 8m,
            Maximum = 32m
        };

        [VisualStudioContribution]
        internal static Setting.String SelectedTabTextColor { get; } = new Setting.String(
            "selectedTabTextColor",
            "Selected — color",
            CustomizationCategory,
            "#FFFFFFFF"
        );

        [VisualStudioContribution]
        internal static Setting.Boolean SelectedTabTextBold { get; } = new Setting.Boolean(
            "selectedTabTextBold",
            "Selected — bold",
            CustomizationCategory,
            true
        );

        [VisualStudioContribution]
        internal static Setting.Boolean SelectedTabTextItalic { get; } = new Setting.Boolean(
            "selectedTabTextItalic",
            "Selected — italic",
            CustomizationCategory,
            false
        );

        [VisualStudioContribution]
        internal static Setting.Decimal SelectedTabTextSize { get; } = new Setting.Decimal(
            "selectedTabTextSize",
            "Selected — size",
            CustomizationCategory,
            12m
        ) {
            Minimum = 8m,
            Maximum = 32m
        };

        [VisualStudioContribution]
        internal static Setting.String ActiveTabTextColor { get; } = new Setting.String(
            "activeTabTextColor",
            "Active — color",
            CustomizationCategory,
            "#FFFFFFFF"
        );

        [VisualStudioContribution]
        internal static Setting.Boolean ActiveTabTextBold { get; } = new Setting.Boolean(
            "activeTabTextBold",
            "Active — bold",
            CustomizationCategory,
            true
        );

        [VisualStudioContribution]
        internal static Setting.Boolean ActiveTabTextItalic { get; } = new Setting.Boolean(
            "activeTabTextItalic",
            "Active — italic",
            CustomizationCategory,
            false
        );

        [VisualStudioContribution]
        internal static Setting.Decimal ActiveTabTextSize { get; } = new Setting.Decimal(
            "activeTabTextSize",
            "Active — size",
            CustomizationCategory,
            12m
        ) {
            Minimum = 8m,
            Maximum = 32m
        };
    }
}
