using System;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace TabsManagerExtension.VsShell.TextEditor {
    [Export(typeof(IWpfTextViewCreationListener))]
    [ContentType("any")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    public class TextEditorFocusTracker : IWpfTextViewCreationListener {
        [Import]
        internal ITextDocumentFactoryService TextDocumentFactoryService { get; set; }

        public static event Action<string> OnTextEditorDocumentChanged;


        public void TextViewCreated(IWpfTextView textView) {
            textView.GotAggregateFocus += (sender, e) => {
                //Helpers.Diagnostic.Logger.LogDebug($"FocusGot [TextEditor]");
                Helpers.FocusWatcher.NotifyFocusGot("TextEditor");
            };

            textView.LostAggregateFocus += (sender, e) => {
                //Helpers.Diagnostic.Logger.LogDebug($"FocusLost [TextEditor]");
                Helpers.FocusWatcher.NotifyFocusLost("TextEditor");
            };
        }
    }
}