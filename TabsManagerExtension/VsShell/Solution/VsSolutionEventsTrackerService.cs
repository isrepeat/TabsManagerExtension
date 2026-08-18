using System;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Utilities;
using EnvDTE;


namespace TabsManagerExtension.VsShell.Solution.Services {
    /// <summary>
    /// Сервис отслеживания загрузки и выгрузки проектов через IVsSolutionEvents.
    /// </summary>
    public sealed class VsSolutionEventsTrackerService :
        TabsManagerExtension.Services.SingletonServiceBase<VsSolutionEventsTrackerService>,
        TabsManagerExtension.Services.IExtensionService,
        IVsSolutionEvents,
        IVsSolutionLoadEvents {

        public event Action<_EventArgs.ProjectHierarchyChangedEventArgs>? ProjectLoaded;
        public event Action<_EventArgs.ProjectHierarchyChangedEventArgs>? ProjectUnloaded;
        public event Action? BackgroundSolutionLoadCompleted;
        public event Action? SolutionHierarchyActivity;

        public bool IsBackgroundSolutionLoadCompleted { get; private set; }

        private IVsSolution? _vsSolution;
        private uint _cookie;
        private uint _solutionLoadCookie;

        public VsSolutionEventsTrackerService() { }

        //
        // IExtensionService
        //
        public IReadOnlyList<Type> DependsOn() {
            return Array.Empty<Type>();
        }

        public void Initialize() {
            ThreadHelper.ThrowIfNotOnUIThread();

            _vsSolution = PackageServices.TryGetVsSolution();
            if (_vsSolution == null) {
                throw new InvalidOperationException("Cannot get SVsSolution");
            }

            ErrorHandler.ThrowOnFailure(_vsSolution.AdviseSolutionEvents(this, out _cookie));

            if (_vsSolution is IVsSolution8 solution8) {
                Guid eventsGuid = typeof(IVsSolutionLoadEvents).GUID;
                ErrorHandler.ThrowOnFailure(solution8.AdviseSolutionEventsEx(ref eventsGuid, this, out _solutionLoadCookie));
            }

            Helpers.Diagnostic.Logger.LogDebug("[VsSolutionEventsTrackerService] Initialized.");
        }

        public void Shutdown() {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_vsSolution != null && _cookie != 0) {
                _vsSolution.UnadviseSolutionEvents(_cookie);
                _cookie = 0;
            }

            if (_vsSolution != null && _solutionLoadCookie != 0) {
                _vsSolution.UnadviseSolutionEvents(_solutionLoadCookie);
                _solutionLoadCookie = 0;
            }

            ClearInstance();
            Helpers.Diagnostic.Logger.LogDebug("[VsSolutionEventsTrackerService] Disposed.");
        }

        //
        // IVsSolutionEvents
        //
        /// <summary>
        /// Этот метод вызывается Visual Studio после открытия (добавления) проекта в Solution.
        /// <para>
        /// В отличие от OnAfterLoadProject, который вызывается только когда проект
        /// "переходит" из IVsStubHierarchy в IVsRealHierarchy (deferred load),
        /// OnAfterOpenProject вызывается в любых случаях:
        /// - при первом открытии проекта,
        /// - при загрузке ранее выгруженного проекта (если не было stub),
        /// - при открытии решения, содержащего этот проект,
        /// - при LoadProject из кода.
        /// </para>
        /// </summary>
        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (fAdded != 0) {
                var dteProject = Utils.EnvDteUtils.GetDteProjectFromHierarchy(pHierarchy);
                Helpers.Diagnostic.Logger.LogDebug($"[VsSolutionEventsTrackerService] OnAfterOpenProject(): {dteProject?.UniqueName}");

                // Visual Studio присылает это событие и для виртуального проекта <MiscFiles>,
                // когда пользователь открывает произвольный файл через File -> Open -> File.
                // Такая hierarchy не входит в solution и не должна попадать в анализаторы проектов.
                if (dteProject == null || Utils.EnvDteUtils.IsMiscProject(dteProject)) {
                    Helpers.Diagnostic.Logger.LogDebug("[VsSolutionEventsTrackerService] Skip non-solution project hierarchy.");
                    return VSConstants.S_OK;
                }

                var newHierarchyItemEntry = Hierarchy.HierarchyItemEntry.CreateWithState<Hierarchy.RealHierarchyItem>(
                    new Hierarchy.HierarchyItemMultiStateElement(
                        pHierarchy,
                        VSConstants.VSITEMID_ROOT
                        ));

                this.ProjectLoaded?.Invoke(new _EventArgs.ProjectHierarchyChangedEventArgs(
                    null, // oldHierarchyItemEntry отсутствует, т.к. проект не переходил из stubHierarchy, а сразу загружен в Solution.
                    newHierarchyItemEntry
                ));
                this.SolutionHierarchyActivity?.Invoke();
            }
            return VSConstants.S_OK;
        }

        public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel) {
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved) {
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy) {
            this.SolutionHierarchyActivity?.Invoke();
            return VSConstants.S_OK;
        }

        public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel) {
            return VSConstants.S_OK;
        }

        public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dteProject = Utils.EnvDteUtils.GetDteProjectFromHierarchy(pRealHierarchy);
            Helpers.Diagnostic.Logger.LogDebug($"[VsSolutionEventsTrackerService] OnBeforeUnloadProject(): {dteProject?.UniqueName}");

            var oldHierarchyItemEntry = Hierarchy.HierarchyItemEntry.CreateWithState<Hierarchy.RealHierarchyItem>(
                new Hierarchy.HierarchyItemMultiStateElement(
                    pRealHierarchy,
                    VSConstants.VSITEMID_ROOT
                    ));

            var newHierarchyItemEntry = Hierarchy.HierarchyItemEntry.CreateWithState<Hierarchy.StubHierarchyItem>(
                new Hierarchy.HierarchyItemMultiStateElement(
                    pStubHierarchy,
                    VSConstants.VSITEMID_ROOT
                    ));

            this.ProjectUnloaded?.Invoke(new _EventArgs.ProjectHierarchyChangedEventArgs(
                oldHierarchyItemEntry,
                newHierarchyItemEntry
                ));

            return VSConstants.S_OK;
        }

        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution) {
            return VSConstants.S_OK;
        }

        public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel) {
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseSolution(object pUnkReserved) {
            return VSConstants.S_OK;
        }

        public int OnAfterCloseSolution(object pUnkReserved) {
            this.IsBackgroundSolutionLoadCompleted = false;
            return VSConstants.S_OK;
        }

        //
        // IVsSolutionLoadEvents
        //
        public int OnBeforeOpenSolution(string pszSolutionFilename) {
            this.IsBackgroundSolutionLoadCompleted = false;
            return VSConstants.S_OK;
        }

        public int OnBeforeBackgroundSolutionLoadBegins() {
            this.IsBackgroundSolutionLoadCompleted = false;
            return VSConstants.S_OK;
        }

        public int OnQueryBackgroundLoadProjectBatch(out bool pfShouldDelayLoadToNextIdle) {
            pfShouldDelayLoadToNextIdle = false;
            return VSConstants.S_OK;
        }

        public int OnBeforeLoadProjectBatch(bool fIsBackgroundIdleBatch) {
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProjectBatch(bool fIsBackgroundIdleBatch) {
            this.SolutionHierarchyActivity?.Invoke();
            return VSConstants.S_OK;
        }

        public int OnAfterBackgroundSolutionLoadComplete() {
            this.IsBackgroundSolutionLoadCompleted = true;
            Helpers.Diagnostic.Logger.LogDebug("[VsSolutionEventsTrackerService] Background solution load completed.");
            this.BackgroundSolutionLoadCompleted?.Invoke();
            return VSConstants.S_OK;
        }
    }
}
