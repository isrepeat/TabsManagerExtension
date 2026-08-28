using System.Collections.Generic;

namespace TabsManagerExtension.Settings {
    internal sealed class TabsManagerSettings {
        public bool AutoLoadCustomTabs { get; set; } = true;
        public bool IsLoggingEnabled { get; set; } = true;
        public string LoggingSessionMode { get; set; } = "LastSession";
        public double TabsScaleFactor { get; set; } = 1.0;
        public string AnchorSectionPattern { get; set; } = TabsManagerSettingsService.DefaultAnchorSectionPattern;
        public string AnchorSubsectionPattern { get; set; } = TabsManagerSettingsService.DefaultAnchorSubsectionPattern;
        public string ActiveSettingsSection { get; set; } = "main";
        public bool ShowTabsToggleToolbarButton { get; set; } = true;
        public bool ShowStandardTabsLayoutToolbarButton { get; set; } = true;
        public bool IsTabEditMode { get; set; }
        public List<string> OpenToolWindowIds { get; set; } = new List<string>();
        public string? ActiveToolWindowId { get; set; }
        public Dictionary<string, string> Appearance { get; set; } = new Dictionary<string, string>();
    }
}
