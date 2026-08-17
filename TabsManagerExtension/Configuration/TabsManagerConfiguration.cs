using System.Collections.Generic;


namespace TabsManagerExtension.Configuration {
    internal sealed class TabsManagerConfiguration {
        public int Version { get; set; } = 1;
        public bool AutoLoadCustomTabs { get; set; }
        public List<string> OpenToolWindowIds { get; set; } = new List<string>();
    }
}
