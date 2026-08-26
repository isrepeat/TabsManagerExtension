using System;
using System.Collections.Generic;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace TabsManagerExtension.VsShell.Services {
    public sealed class VsIDEStateFlagsTrackerService :
        VsShell.Services.VsSelectionEventsServiceBase<VsIDEStateFlagsTrackerService>,
        TabsManagerExtension.Services.IExtensionService {
        
        public static readonly Guid SolutionExistsGuid = new Guid(UIContextGuids80.SolutionExists);

        public readonly Helpers.Events.Action<Guid, bool> IDEStateFlagsChanged = new();

        private readonly Dictionary<Guid, bool> _contextStateMap = new();
        private readonly Dictionary<uint, Guid> _contextCookiesMap = new();

        public VsIDEStateFlagsTrackerService() { }

        //
        // IExtensionService
        //
        public IReadOnlyList<Type> DependsOn() {
            return Array.Empty<Type>();
        }

        public void Initialize() {
            ThreadHelper.ThrowIfNotOnUIThread();

            IEnumerable<Guid> trackedContextGuids = new List<Guid> {
                SolutionExistsGuid,
            };

            foreach (var guid in trackedContextGuids) {
                var contextGuid = guid; // make copy for pass by ref

                int hr = PackageServices.VsMonitorSelection.GetCmdUIContextCookie(ref contextGuid, out uint cookie);
                ErrorHandler.ThrowOnFailure(hr);

                _contextCookiesMap[cookie] = contextGuid;

                // Проверяем состояние прямо сейчас для каждого контекста
                hr = PackageServices.VsMonitorSelection.IsCmdUIContextActive(cookie, out int pfActive);
                if (ErrorHandler.Succeeded(hr)) {
                    this.HandleContextState(contextGuid, pfActive != 0);
                }
            }

            Helpers.Diagnostic.Logger.LogDebug("[VsIDEStateFlagsTrackerService] Initialized.");
        }

        public void Shutdown() {
            ThreadHelper.ThrowIfNotOnUIThread();

            base.Dispose();

            ClearInstance();
            Helpers.Diagnostic.Logger.LogDebug("[VsIDEStateFlagsTrackerService] Disposed.");
        }

        //
        // VsSelectionEventsServiceBase
        //
        public override int OnCmdUIContextChanged(uint dwCmdUICookie, int fActive) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_contextCookiesMap.TryGetValue(dwCmdUICookie, out Guid guid)) {
                this.HandleContextState(guid, fActive != 0);
            }

            return VSConstants.S_OK;
        }


        //
        // Api
        //
        public bool IsContextActive(Guid contextGuid) {
            return _contextStateMap.TryGetValue(contextGuid, out bool isActive) && isActive;
        }

        //
        // Internal logic
        //
        private void HandleContextState(Guid contextGuid, bool isActive) {
            _contextStateMap[contextGuid] = isActive;

            this.IDEStateFlagsChanged.Invoke(contextGuid, isActive);
        }
    }
}
