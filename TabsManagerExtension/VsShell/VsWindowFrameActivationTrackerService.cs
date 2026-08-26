using System;
using System.Collections.Generic;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace TabsManagerExtension.VsShell.Solution.Services {
    /// <summary>
    /// Отслеживает активацию окон (IVsWindowFrame), когда активируется новое окно VS 
    /// (будь то редактор, Solution Explorer, Output и т.д.)
    /// </summary>
    public sealed class VsWindowFrameActivationTrackerService :
         VsShell.Services.VsSelectionEventsServiceBase<VsWindowFrameActivationTrackerService>,
         TabsManagerExtension.Services.IExtensionService {

        public event Action<IVsWindowFrame>? VsWindowFrameActivated;

        public VsWindowFrameActivationTrackerService() { }

        //
        // IExtensionService
        //
        public IReadOnlyList<Type> DependsOn() {
            return Array.Empty<Type>();
        }

        public void Initialize() {
            ThreadHelper.ThrowIfNotOnUIThread();
            Helpers.Diagnostic.Logger.LogDebug("[VsWindowFrameActivationTrackerService] Initialized.");
        }

        public void Shutdown() {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            base.Dispose();
            
            ClearInstance();
            Helpers.Diagnostic.Logger.LogDebug("[VsWindowFrameActivationTrackerService] Disposed.");
        }

        //
        // VsSelectionEventsServiceBase
        //
        public override int OnElementValueChanged(uint elementid, object oldValue, object newValue) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (elementid == (uint)VSConstants.VSSELELEMID.SEID_WindowFrame && newValue is IVsWindowFrame frame) {
                this.VsWindowFrameActivated?.Invoke(frame);
            }

            return VSConstants.S_OK;
        }
    }
}
