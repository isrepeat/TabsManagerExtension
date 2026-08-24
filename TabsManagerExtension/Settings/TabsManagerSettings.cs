using System.Collections.Generic;


namespace TabsManagerExtension.Settings {
    internal sealed class TabsManagerSettings {
        public int Version { get; set; } = 2;
        public bool AutoLoadCustomTabs { get; set; } = true;
        public double TabsScaleFactor { get; set; } = 1.0;
        public string AnchorSectionPattern { get; set; } = TabsManagerSettingsService.DefaultAnchorSectionPattern;
        public string AnchorSubsectionPattern { get; set; } = TabsManagerSettingsService.DefaultAnchorSubsectionPattern;
        public string ActiveSettingsSection { get; set; } = "main";
        public bool ShowTabsToggleToolbarButton { get; set; } = true;
        public bool ShowStandardTabsLayoutToolbarButton { get; set; } = true;
        public List<string> OpenToolWindowIds { get; set; } = new List<string>();
        public string? ActiveToolWindowId { get; set; }
        public Dictionary<string, string> Appearance { get; set; } = new Dictionary<string, string>();
    }
}
