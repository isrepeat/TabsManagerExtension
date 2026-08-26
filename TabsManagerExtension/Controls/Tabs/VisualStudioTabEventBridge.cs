using System;
using Microsoft.VisualStudio.Shell.Interop;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Адаптирует DTE и VS tracker events к событиям подсистемы вкладок.</summary>
    internal sealed class VisualStudioTabEventBridge : IDisposable {
        private readonly EnvDTE80.DTE2 _dte;

        private EnvDTE.WindowEvents? _windowEvents;
        private EnvDTE.DocumentEvents? _documentEvents;

        public event Action<EnvDTE.Document>? DocumentOpened;
        public event Action<EnvDTE.Document>? DocumentSaved;
        public event Action<EnvDTE.Document>? DocumentClosing;
        public event Action<EnvDTE.Window, EnvDTE.Window>? WindowActivated;
        public event Action<EnvDTE.Window>? WindowClosing;
        public event Action<VsShell._EventArgs.DocumentNavigationEventArgs>? DocumentActivatedExternally;
        public event Action<IVsWindowFrame>? WindowFrameActivated;
        public event Action<VsShell._EventArgs.DocumentNavigationEventArgs>? TextEditorNavigatedToDocument;

        public VisualStudioTabEventBridge(EnvDTE80.DTE2 dte) {
            _dte = dte;
        }

        public void Start() {
            if (_documentEvents != null) {
                return;
            }

            _documentEvents = _dte.Events.DocumentEvents;
            _documentEvents.DocumentOpened += this.OnDocumentOpened;
            _documentEvents.DocumentSaved += this.OnDocumentSaved;
            _documentEvents.DocumentClosing += this.OnDocumentClosing;

            _windowEvents = _dte.Events.WindowEvents;
            _windowEvents.WindowActivated += this.OnWindowActivated;
            _windowEvents.WindowClosing += this.OnWindowClosing;

            VsShell.Document.Services.VsDocumentActivationTrackerService.Instance.OnDocumentActivated += this.OnDocumentActivatedExternally;
            VsShell.Solution.Services.VsWindowFrameActivationTrackerService.Instance.VsWindowFrameActivated += this.OnWindowFrameActivated;
            VsShell.TextEditor.Services.TextEditorFileNavigationCommandFilterService.Instance.OnNavigatedToDocument += this.OnTextEditorNavigatedToDocument;
        }

        public void Stop() {
            if (_documentEvents == null) {
                return;
            }

            VsShell.TextEditor.Services.TextEditorFileNavigationCommandFilterService.Instance.OnNavigatedToDocument -= this.OnTextEditorNavigatedToDocument;
            VsShell.Solution.Services.VsWindowFrameActivationTrackerService.Instance.VsWindowFrameActivated -= this.OnWindowFrameActivated;
            VsShell.Document.Services.VsDocumentActivationTrackerService.Instance.OnDocumentActivated -= this.OnDocumentActivatedExternally;

            if (_windowEvents != null) {
                _windowEvents.WindowClosing -= this.OnWindowClosing;
                _windowEvents.WindowActivated -= this.OnWindowActivated;
                _windowEvents = null;
            }

            _documentEvents.DocumentClosing -= this.OnDocumentClosing;
            _documentEvents.DocumentSaved -= this.OnDocumentSaved;
            _documentEvents.DocumentOpened -= this.OnDocumentOpened;
            _documentEvents = null;
        }

        public void Dispose() {
            this.Stop();
        }

        private void OnDocumentOpened(EnvDTE.Document document) {
            this.DocumentOpened?.Invoke(document);
        }

        private void OnDocumentSaved(EnvDTE.Document document) {
            this.DocumentSaved?.Invoke(document);
        }

        private void OnDocumentClosing(EnvDTE.Document document) {
            this.DocumentClosing?.Invoke(document);
        }

        private void OnWindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window lostFocus) {
            this.WindowActivated?.Invoke(gotFocus, lostFocus);
        }

        private void OnWindowClosing(EnvDTE.Window window) {
            this.WindowClosing?.Invoke(window);
        }

        private void OnDocumentActivatedExternally(VsShell._EventArgs.DocumentNavigationEventArgs eventArgs) {
            this.DocumentActivatedExternally?.Invoke(eventArgs);
        }

        private void OnWindowFrameActivated(IVsWindowFrame windowFrame) {
            this.WindowFrameActivated?.Invoke(windowFrame);
        }

        private void OnTextEditorNavigatedToDocument(VsShell._EventArgs.DocumentNavigationEventArgs eventArgs) {
            this.TextEditorNavigatedToDocument?.Invoke(eventArgs);
        }
    }
}
