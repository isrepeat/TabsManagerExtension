#pragma warning disable VSEXTPREVIEW_SETTINGS

using Microsoft.VisualStudio.Extensibility;
using Microsoft.VisualStudio.Extensibility.Settings;

namespace TabsManagerExtension.Configuration {
    internal static class TabsManagerSettingDefinitions {
        [VisualStudioContribution]
        internal static SettingCategory TabsManagerCategory { get; } = new SettingCategory(
            "tabsManager",
            "Tabs Manager"
        ) {
            Description = "Settings for the Tabs Manager extension."
        };

        [VisualStudioContribution]
        internal static Setting.Boolean AutoLoadCustomTabs { get; } = new Setting.Boolean(
            "autoLoadCustomTabs",
            "Enable custom tabs when Visual Studio starts",
            TabsManagerCategory,
            defaultValue: true
        ) {
            Description = "Automatically replaces the standard tab list with Tabs Manager after Visual Studio starts."
        };

        [VisualStudioContribution]
        internal static Setting.Decimal TabsScaleFactor { get; } = new Setting.Decimal(
            "tabsScaleFactor",
            "Tab compactness",
            TabsManagerCategory,
            defaultValue: 1.0m
        ) {
            Description = "Controls the size and spacing of tabs.",
            Minimum = 0.5m,
            Maximum = 1.5m
        };
    }
}
