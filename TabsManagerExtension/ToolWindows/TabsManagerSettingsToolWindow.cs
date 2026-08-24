using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace TabsManagerExtension.ToolWindows {
    [Guid("66ee240d-e51c-4710-a0c2-96c82ef6e7c8")]
    public sealed class TabsManagerSettingsToolWindow : ToolWindowPane {
        public TabsManagerSettingsToolWindow() : base(null) {
            this.Caption = "Tabs Manager Settings";
            this.Content = new Controls.TabsManagerSettingsControl();
        }
    }
}
