using System.Collections.Generic;


namespace TabsManagerExtension.Configuration {
    internal sealed class TabsManagerConfiguration {
        public int Version { get; set; } = 1;
        public bool AutoLoadCustomTabs { get; set; }
        public double TabsScaleFactor { get; set; } = 1.0;
        public List<string> OpenToolWindowIds { get; set; } = new List<string>();
    }
}
