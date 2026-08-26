using System;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;

namespace TabsManagerExtension.VsShell {
    public abstract class OleCommandFilterBase : IOleCommandTarget {
        private IOleCommandTarget? _next;

        //
        // IOleCommandTarget
        //
        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText) {
            return _next?.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText) ?? VSConstants.E_FAIL;
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut) {
            if (this.TryHandleCommand(pguidCmdGroup, nCmdID, pvaIn)) {
                this.OnCommandIntercepted(pguidCmdGroup, nCmdID, pvaIn);
                return VSConstants.S_OK;
            }

            this.OnCommandPassedThrough(pguidCmdGroup, nCmdID, pvaIn);
            return _next?.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut) ?? VSConstants.E_FAIL;
        }

        //
        // Api
        //
        public void SetNext(IOleCommandTarget next) {
            _next = next;
        }

        protected abstract bool TryHandleCommand(Guid cmdGroup, uint cmdId, IntPtr inputArgument);

        protected virtual void OnCommandIntercepted(Guid cmdGroup, uint cmdId, IntPtr inputArgument) { }

        protected virtual void OnCommandPassedThrough(Guid cmdGroup, uint cmdId, IntPtr inputArgument) { }
    }
}
