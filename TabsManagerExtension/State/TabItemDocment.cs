using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace TabsManagerExtension.State.Document {
    public interface IActivatableTab {
        void Activate();
    }

    public class TabItemDocument : TabItemBase, IActivatableTab {
        public VsShell.Document.ShellDocument ShellDocument { get; private set; }
        public DocumentProjectReferencesInfo DocumentProjectReferencesInfo { get; }
        //public VsShell.Project.ProjectNode ProjectNodeContext { get; set; }

        public TabItemDocument(VsShell.Document.ShellDocument shellDocument) {
            ThreadHelper.ThrowIfNotOnUIThread();

            base.Caption = shellDocument.Document.Name;
            base.FullName = shellDocument.Document.FullName;
            this.ShellDocument = shellDocument;
            this.DocumentProjectReferencesInfo = new DocumentProjectReferencesInfo(base.FullName);
        }

        public TabItemDocument(EnvDTE.Document document)
            : this(new VsShell.Document.ShellDocument(document)) {
        }

        public void Activate() {
            ThreadHelper.ThrowIfNotOnUIThread();

            try {
                this.ShellDocument.Document?.Activate();
            }
            catch (COMException ex) {
                Helpers.Diagnostic.Logger.LogWarning($"Failed to activate document '{this.Caption}': {ex.Message}");
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Unexpected error activating document '{this.Caption}': {ex.Message}");
            }
        }


        public override string ToString() {
            //return $"TabItemDocument(FullName='{this.FullName}', ProjectCtx='{this.ProjectNodeContext}')";
            return $"TabItemDocument(FullName='{this.FullName}'')";
        }
    }
}
