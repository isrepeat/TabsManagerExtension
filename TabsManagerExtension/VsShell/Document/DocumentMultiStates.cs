using System;
using System.Linq;
using System.ComponentModel;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Helpers.Attributes;
using System.Threading.Tasks;
using System.Collections.Generic;
using TabsManagerExtension.VsShell.Hierarchy;


namespace TabsManagerExtension.VsShell.Document {
    public abstract class DocumentMultiStateElementBase :
        Helpers.MultiState.MultiStateContainer<
            _Details.DocumentCommonState,
            Document,
            InvalidatedDocument> {

        protected DocumentMultiStateElementBase(_Details.DocumentCommonState commonState)
            : base(commonState) {
        }

        protected DocumentMultiStateElementBase(
            _Details.DocumentCommonState commonState,
            Func<_Details.DocumentCommonState, Document> factoryA,
            Func<_Details.DocumentCommonState, InvalidatedDocument> factoryB)
            : base(commonState, factoryA, factoryB) {
        }
    }



    public class DocumentMultiStateElement : DocumentMultiStateElementBase {
        public DocumentMultiStateElement(
            Project.ProjectCommonStateViewModel projectBaseViewModel,
            Hierarchy.HierarchyItemEntry hierarchyItemEntry
            ) : base(new _Details.DocumentCommonState(projectBaseViewModel, hierarchyItemEntry)) {
        }
    }


    public class SharedItemMultiStateElement : DocumentMultiStateElementBase {
        public SharedItemMultiStateElement(
            Project.ProjectCommonStateViewModel projectBaseViewModel,
            Hierarchy.HierarchyItemEntry hierarchyItemEntry
            ) : base(
                new _Details.DocumentCommonState(projectBaseViewModel, hierarchyItemEntry),
                commonState => new SharedItem(commonState),
                commonState => new InvalidatedDocument(commonState)
                ) {
        }
    }


    public class ExternalIncludeMultiStateElement : DocumentMultiStateElementBase {
        public ExternalIncludeMultiStateElement(
            Project.ProjectCommonStateViewModel projectBaseViewModel,
            Hierarchy.HierarchyItemEntry hierarchyItemEntry
            ) : base(
                new _Details.DocumentCommonState(projectBaseViewModel, hierarchyItemEntry),
                commonState => new ExternalInclude(commonState),
                commonState => new InvalidatedDocument(commonState)
                ) {
        }
    }



    public partial class Document :
        DocumentCommonStateViewModel,
        Helpers.MultiState.IMultiStateElement {

        [ObservableProperty(AccessMarker.Get, AccessMarker.PrivateSet)]
        private bool _isOppenedWithProjectContext = false;

        public Document(_Details.DocumentCommonState commonState) : base(commonState) {
        }

        public void OnStateEnabled(Helpers._EventArgs.MultiStateElementEnabledEventArgs e) {
            if (e.PreviousState is Helpers.MultiState.UnknownMultiStateElement) {
                Helpers.ThrowableAssert.Require(base.CommonState.HierarchyItemEntry.IsRealHierarchy);
            }
        }

        public void OnStateDisabled(Helpers._EventArgs.MultiStateElementDisabledEventArgs e) {
        }

        public override string ToString() {
            return $"<Document> ({base.CommonState.ToStringCore()})";
        }

        protected override void OnCommonStatePropertyChanged(object? sender, PropertyChangedEventArgs e) {
            base.OnCommonStatePropertyChanged(sender, e);
        }

        protected void OpenWithProjectContext(bool restoreActiveDocument = true) {
            ThreadHelper.ThrowIfNotOnUIThread();

            Helpers.ThrowableAssert.Require(!base.CommonState.IsDisposed);
            Helpers.ThrowableAssert.Require(base.CommonState.HierarchyItemEntry.MultiState.Current is Hierarchy.RealHierarchyItem);
            Helpers.ThrowableAssert.Require(base.ProjectBaseViewModel is Project.LoadedProject);

            var ownProject = base.ProjectBaseViewModel as Project.LoadedProject;

            // Сохраняем активный документ до всех действий.
            var activeDocumentBefore = PackageServices.Dte2.ActiveDocument;

            // Попытка найти первый cpp/h файл проекта,
            // чтобы открыть его и "переключить" контекст редактора на нужный проект.
            // Это нужно для того, чтобы при открытии внешнего include файла
            // Visual Studio знала, что контекстом открытия является именно этот проект.
            var includeDependencyAnalyzer = VsShell.Solution.Services.IncludeDependencyAnalyzerService.Instance;
            var allTransitiveIncludingFiles = includeDependencyAnalyzer
                .GetTransitiveFilesIncludersByIncludePath(base.HierarchyItemEntry.BaseViewModel.FilePath);

            var currentProjectTransitiveIncludingFiles = allTransitiveIncludingFiles
                .Where(sf => sf.LoadedProject.Equals(ownProject))
                .ToList();

            // Нужно открывать именно .cpp файл, который реально включает наш include.
            // Закрытый файл предпочтительнее: после короткой активации мы сразу закроем его,
            // поэтому пользователь почти не увидит служебную смену editor frame.
            var openDocumentPaths = PackageServices.Dte2.Documents
                .Cast<EnvDTE.Document>()
                .Select(d => d.FullName)
                .Where(path => !string.IsNullOrEmpty(path))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var solutionHierarchyAnalyzer = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance;
            var contextSwitchCandidates = currentProjectTransitiveIncludingFiles
                .Where(sf => string.Equals(
                    System.IO.Path.GetExtension(sf.FilePath),
                    ".cpp",
                    StringComparison.OrdinalIgnoreCase
                ))
                .Select(sf => new {
                    SourceFile = sf,
                    Document = solutionHierarchyAnalyzer.SourcesRepresentationsTable
                        .GetDocumentByProjectAndDocumentPath(base.ProjectBaseViewModel, sf.FilePath)
                })
                .Where(candidate => candidate.Document != null)
                .ToList();

            var contextSwitchCandidate = contextSwitchCandidates
                .FirstOrDefault(candidate => !openDocumentPaths.Contains(candidate.SourceFile.FilePath))
                ?? contextSwitchCandidates.FirstOrDefault();

            if (contextSwitchCandidate == null) {
                Helpers.Diagnostic.Logger.LogDebug($"[Document.OpenWithProjectContext] No usable .cpp representation was found in project '{ownProject.UniqueName}'.");
                return;
            }

            var contextSwitchDocument = contextSwitchCandidate.Document;
            var contextSwitchDocumentProject = contextSwitchDocument.BaseViewModel.ProjectBaseViewModel as Project.LoadedProject;
            if (contextSwitchDocumentProject == null) {
                Helpers.Diagnostic.Logger.LogDebug($"[Document.OpenWithProjectContext] Source representation '{contextSwitchCandidate.SourceFile.FilePath}' has no loaded project.");
                return;
            }

            var contextSwitchDocumentHierarchyItem = contextSwitchDocument.BaseViewModel.HierarchyItemEntry.BaseViewModel as Hierarchy.RealHierarchyItem;
            if (contextSwitchDocumentHierarchyItem == null) {
                Helpers.Diagnostic.Logger.LogDebug($"[Document.OpenWithProjectContext] Source representation '{contextSwitchCandidate.SourceFile.FilePath}' has no real hierarchy item.");
                return;
            }

            bool needCloseContextSwitchDocumentNode =
                !openDocumentPaths.Contains(contextSwitchDocumentHierarchyItem.FilePath);

            Helpers.Diagnostic.Logger.LogDebug($"[Document.OpenWithProjectContext] Context switch source '{contextSwitchDocumentHierarchyItem.FilePath}' was {(needCloseContextSwitchDocumentNode ? "closed" : "already open")} before activation.");

            int hr = VSConstants.S_OK;

            // Открываем файл в контексте проекта.
            hr = Utils.VsHierarchyUtils.ClickOnSolutionHierarchyItem(
                ownProject.ProjectHierarchy.VsRealHierarchy,
                base.HierarchyItemEntry.BaseViewModel.ItemId);
            ErrorHandler.ThrowOnFailure(hr);

            // Переключаемся на файл который включает наш файл (для смены activeDocumentFrame)
            // иначе IntelliSense не подхватит контекст.
            hr = Utils.VsHierarchyUtils.ClickOnSolutionHierarchyItem(
                contextSwitchDocumentProject.ProjectHierarchy.VsRealHierarchy,
                contextSwitchDocumentHierarchyItem.ItemId);
            ErrorHandler.ThrowOnFailure(hr);
            

            // Закрываем временный файл переключения контекста.
            if (needCloseContextSwitchDocumentNode) {
                var doc = PackageServices.Dte2.Documents.Cast<EnvDTE.Document>()
                    .FirstOrDefault(d =>
                        string.Equals(
                            d.FullName,
                            contextSwitchDocumentHierarchyItem.FilePath,
                            StringComparison.OrdinalIgnoreCase)
                        );

                doc?.Close(EnvDTE.vsSaveChanges.vsSaveChangesNo);
            }

            this.IsOppenedWithProjectContext = true;

            if (restoreActiveDocument) {
                // Возвращаем активным предыдущий документ.
                VsixThreadHelper.RunOnUiThread(async () => {
                    await Task.Delay(20);
                    activeDocumentBefore?.Activate();
                });
            }
        }
    }



    //
    // SharedItem
    //
    public sealed class SharedItem : Document {
        public SharedItem(_Details.DocumentCommonState commonState) : base(commonState) {
        }

        public new void OpenWithProjectContext() {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (base.ProjectBaseViewModel is not Project.LoadedProject targetProject) {
                Helpers.Diagnostic.Logger.LogError("[SharedItem.OpenWithProjectContext] Target project is not loaded.");
                return;
            }

            var solutionHierarchyAnalyzer = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance;
            var sharedItemProjects = solutionHierarchyAnalyzer.SharedItemsRepresentationsTable
                .GetProjectsByDocumentPath(base.HierarchyItemEntry.BaseViewModel.FilePath);

            var sharedProject = sharedItemProjects
                .Select(p => p.MultiState.Current)
                .OfType<Project.LoadedProject>()
                .FirstOrDefault(p => p.IsSharedProject);

            if (sharedProject == null) {
                Helpers.Diagnostic.Logger.LogError($"[SharedItem.OpenWithProjectContext] Shared project was not found for '{base.HierarchyItemEntry.BaseViewModel.FilePath}'.");
                return;
            }

            var sharedHierarchy = sharedProject.ProjectHierarchy.VsRealHierarchy;
            var targetHierarchy = targetProject.ProjectHierarchy.VsRealHierarchy;
            int hr = sharedHierarchy.SetProperty(
                VSConstants.VSITEMID_ROOT,
                (int)__VSHPROPID7.VSHPROPID_SharedItemContextHierarchy,
                targetHierarchy
            );

            if (ErrorHandler.Failed(hr)) {
                Helpers.Diagnostic.Logger.LogError($"[SharedItem.OpenWithProjectContext] Failed to set context '{targetProject.UniqueName}' on '{sharedProject.UniqueName}': HRESULT=0x{hr:X8}.");
                return;
            }

            sharedHierarchy.GetProperty(
                VSConstants.VSITEMID_ROOT,
                (int)__VSHPROPID7.VSHPROPID_SharedItemContextHierarchy,
                out var actualContextObject
            );

            var actualContextHierarchy = actualContextObject as IVsHierarchy;
            if (actualContextHierarchy == null) {
                Helpers.Diagnostic.Logger.LogError($"[SharedItem.OpenWithProjectContext] Visual Studio returned no active context for '{sharedProject.UniqueName}'.");
                return;
            }

            PackageServices.VsSolution.GetGuidOfProject(targetHierarchy, out var targetProjectGuid);
            PackageServices.VsSolution.GetGuidOfProject(actualContextHierarchy, out var actualContextProjectGuid);

            if (actualContextProjectGuid != targetProjectGuid) {
                Helpers.Diagnostic.Logger.LogError($"[SharedItem.OpenWithProjectContext] Visual Studio did not retain context '{targetProject.UniqueName}'.");
                return;
            }

            Helpers.Diagnostic.Logger.LogDebug($"[SharedItem.OpenWithProjectContext] Context switched: '{sharedProject.UniqueName}' -> '{targetProject.UniqueName}'.");

            var openDocument = PackageServices.Dte2.Documents
                .Cast<EnvDTE.Document>()
                .FirstOrDefault(d => string.Equals(
                    d.FullName,
                    base.HierarchyItemEntry.BaseViewModel.FilePath,
                    StringComparison.OrdinalIgnoreCase
                ));

            int? caretLine = null;
            int? caretOffset = null;

            if (openDocument != null && !openDocument.Saved) {
                Helpers.Diagnostic.Logger.LogError($"[SharedItem.OpenWithProjectContext] Cannot reopen modified document '{openDocument.FullName}'. Save it before switching project context.");
                return;
            }

            if (openDocument?.Selection is EnvDTE.TextSelection textSelection) {
                caretLine = textSelection.ActivePoint.Line;
                caretOffset = textSelection.ActivePoint.LineCharOffset;
            }

            bool TryRetargetOpenDocumentFrame() {
                ThreadHelper.ThrowIfNotOnUIThread();

                PackageServices.VsUIShell.GetDocumentWindowEnum(out var windowFramesEnum);
                var frames = new IVsWindowFrame[1];

                while (windowFramesEnum.Next(1, frames, out uint fetched) == VSConstants.S_OK && fetched == 1) {
                    var frame = frames[0];
                    int pathHr = frame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out var documentPathObject);
                    if (ErrorHandler.Failed(pathHr) ||
                        documentPathObject is not string documentPath ||
                        !string.Equals(documentPath, base.HierarchyItemEntry.BaseViewModel.FilePath, StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    frame.GetProperty((int)__VSFPROPID.VSFPROPID_Hierarchy, out var previousHierarchyObject);
                    frame.GetProperty((int)__VSFPROPID.VSFPROPID_ItemID, out var previousItemIdObject);

                    int hierarchyHr = frame.SetProperty((int)__VSFPROPID.VSFPROPID_Hierarchy, targetHierarchy);
                    int itemIdHr = ErrorHandler.Succeeded(hierarchyHr)
                        ? frame.SetProperty((int)__VSFPROPID.VSFPROPID_ItemID, base.HierarchyItemEntry.BaseViewModel.ItemId)
                        : hierarchyHr;

                    frame.GetProperty((int)__VSFPROPID.VSFPROPID_Hierarchy, out var actualHierarchyObject);
                    frame.GetProperty((int)__VSFPROPID.VSFPROPID_ItemID, out var actualItemIdObject);

                    var actualHierarchy = actualHierarchyObject as IVsHierarchy;
                    PackageServices.VsSolution.GetGuidOfProject(actualHierarchy, out var actualHierarchyGuid);
                    uint actualItemId = Convert.ToUInt32(actualItemIdObject);
                    bool retargeted =
                        ErrorHandler.Succeeded(hierarchyHr) &&
                        ErrorHandler.Succeeded(itemIdHr) &&
                        actualHierarchyGuid == targetProjectGuid &&
                        actualItemId == base.HierarchyItemEntry.BaseViewModel.ItemId;

                    if (retargeted) {
                        Helpers.Diagnostic.Logger.LogDebug($"[SharedItem.OpenWithProjectContext] Existing frame retargeted to '{targetProject.UniqueName}' without reopening.");
                        return true;
                    }

                    if (previousHierarchyObject != null) {
                        frame.SetProperty((int)__VSFPROPID.VSFPROPID_Hierarchy, previousHierarchyObject);
                    }

                    if (previousItemIdObject != null) {
                        frame.SetProperty((int)__VSFPROPID.VSFPROPID_ItemID, previousItemIdObject);
                    }

                    Helpers.Diagnostic.Logger.LogDebug($"[SharedItem.OpenWithProjectContext] Frame retarget was rejected: hierarchy HRESULT=0x{hierarchyHr:X8}, item HRESULT=0x{itemIdHr:X8}. Falling back to reopen.");
                    return false;
                }

                Helpers.Diagnostic.Logger.LogDebug("[SharedItem.OpenWithProjectContext] Open document frame was not found. Falling back to reopen.");
                return false;
            }

            void ReopenDocumentInTargetHierarchy() {
                ThreadHelper.ThrowIfNotOnUIThread();

                var documentToReopen = PackageServices.Dte2.Documents
                    .Cast<EnvDTE.Document>()
                    .FirstOrDefault(d => string.Equals(
                        d.FullName,
                        base.HierarchyItemEntry.BaseViewModel.FilePath,
                        StringComparison.OrdinalIgnoreCase
                    ));

                Helpers.Diagnostic.Logger.LogDebug($"[SharedItem.OpenWithProjectContext] Closing existing document frame before reopening in '{targetProject.UniqueName}'.");
                documentToReopen?.Close(EnvDTE.vsSaveChanges.vsSaveChangesNo);

                int reopenHr = Utils.VsHierarchyUtils.ClickOnSolutionHierarchyItem(
                    targetHierarchy,
                    base.HierarchyItemEntry.BaseViewModel.ItemId
                );

                ErrorHandler.ThrowOnFailure(reopenHr);

                if (caretLine.HasValue && caretOffset.HasValue) {
                    var reopenedDocument = PackageServices.Dte2.Documents
                        .Cast<EnvDTE.Document>()
                        .FirstOrDefault(d => string.Equals(
                            d.FullName,
                            base.HierarchyItemEntry.BaseViewModel.FilePath,
                            StringComparison.OrdinalIgnoreCase
                        ));

                    if (reopenedDocument?.Selection is EnvDTE.TextSelection reopenedSelection) {
                        reopenedSelection.MoveToLineAndOffset(caretLine.Value, caretOffset.Value, false);
                    }
                }
            }

            bool frameRetargeted = TryRetargetOpenDocumentFrame();

            // C++ language service не всегда применяет новый shared-context только по событию
            // VSHPROPID_SharedItemContextHierarchy. Кратковременно активируем включающий .cpp,
            // сразу возвращаем header, а после обработки смены translation unit переоткрываем
            // его в новой hierarchy. Так уже открытый .cpp не остаётся видимым на время ожидания.
            if (!targetProject.IsSharedProject) {
                if (frameRetargeted) {
                    Helpers.Diagnostic.Logger.LogDebug($"[SharedItem.OpenWithProjectContext] Including .cpp activation skipped because the existing frame already accepted context '{targetProject.UniqueName}'.");
                    return;
                }

                base.OpenWithProjectContext(restoreActiveDocument: true);

                VsixThreadHelper.RunOnUiThread(async () => {
                    await Task.Delay(300);
                    ReopenDocumentInTargetHierarchy();
                });

                return;
            }

            // Сам shared-проект владеет файлом, но не имеет включающего translation unit.
            if (!frameRetargeted) {
                ReopenDocumentInTargetHierarchy();
            }

            var reopenedSharedDocument = PackageServices.Dte2.Documents
                .Cast<EnvDTE.Document>()
                .FirstOrDefault(d => string.Equals(
                    d.FullName,
                    base.HierarchyItemEntry.BaseViewModel.FilePath,
                    StringComparison.OrdinalIgnoreCase
                ));

            var frameSwitchDocument = PackageServices.Dte2.Documents
                .Cast<EnvDTE.Document>()
                .FirstOrDefault(d => !string.Equals(
                    d.FullName,
                    base.HierarchyItemEntry.BaseViewModel.FilePath,
                    StringComparison.OrdinalIgnoreCase
                ));

            if (reopenedSharedDocument != null && frameSwitchDocument != null) {
                Helpers.Diagnostic.Logger.LogDebug($"[SharedItem.OpenWithProjectContext] Refreshing owner context through frame '{frameSwitchDocument.Name}'.");
                frameSwitchDocument.Activate();

                VsixThreadHelper.RunOnUiThread(async () => {
                    await Task.Delay(100);
                    reopenedSharedDocument.Activate();
                });
            }
            else {
                Helpers.Diagnostic.Logger.LogDebug("[SharedItem.OpenWithProjectContext] Owner context frame refresh was skipped because no other open document is available.");
            }

            Helpers.Diagnostic.Logger.LogDebug($"[SharedItem.OpenWithProjectContext] Owner context '{targetProject.UniqueName}' does not require an including .cpp activation.");
        }

        public override string ToString() {
            return $"<SharedItem> ({base.CommonState.ToStringCore()})";
        }
    }



    //
    // ExternalInclude
    //
    public sealed class ExternalInclude : Document {
        public ExternalInclude(_Details.DocumentCommonState commonState) : base(commonState) {
        }

        public new void OpenWithProjectContext() {
            base.OpenWithProjectContext();
        }

        public override string ToString() {
            return $"<ExternalInclude> ({base.CommonState.ToStringCore()})";
        }
    }




    public class InvalidatedDocument :
        DocumentCommonStateViewModel,
        Helpers.MultiState.IMultiStateElement {

        private Document? _invalidatedPreviousDocument = null;
        private string _invalidatedPreviousDocumentFilePath;

        public InvalidatedDocument(_Details.DocumentCommonState commonState) : base(commonState) {
        }

        public void OnStateEnabled(Helpers._EventArgs.MultiStateElementEnabledEventArgs e) {
            if (e.PreviousState is Document document) {
                _invalidatedPreviousDocument = document;
                _invalidatedPreviousDocumentFilePath = document.HierarchyItemEntry.BaseViewModel.FilePath;
            }

            if (e.UpdatePackageObj is Project.ProjectCommonStateViewModel projectBaseViewModel) {
                base.CommonState.ProjectBaseViewModel = projectBaseViewModel;
            }

            base.CommonState.HierarchyItemEntry.MultiState.SwitchTo<Hierarchy.InvalidatedHierarchyItem>();
        }

        public void OnStateDisabled(Helpers._EventArgs.MultiStateElementDisabledEventArgs e) {
            Helpers.ThrowableAssert.Unexpected("Switching from InvalidatedDocument is not supported");
        }

        public void OpenWithProjectContext() {
            if (base.ProjectBaseViewModel is Project.UnloadedProject) {
                if (_invalidatedPreviousDocument is SharedItem) {
                    base.ProjectBaseViewModel.SharedItemsChanged.Add(
                        Helpers.Events.Action.Options.UnsubscribeAfterInvoked,
                        this.OnFreshedSharedItemsLoaded);
                }

                Utils.VsHierarchyUtils.ReloadProject(base.ProjectBaseViewModel.ProjectGuid);
                // NOTE: After reload this document will be disposed.
            }
        }

        public override string ToString() {
            return $"<InvalidatedDocument> ({base.CommonState.ToStringCore()})";
        }

        private void OnFreshedSharedItemsLoaded(IReadOnlyList<SharedItemEntry> freshSharedItemEntries) {
            Helpers.ThrowableAssert.Require(freshSharedItemEntries.All(d => d.MultiState.Current is SharedItem));
            Helpers.ThrowableAssert.Require(base.IsDisposed);

            // WORKAROUND:
            // Выполняем в следующей итерации очереди, т.к. OpenWithProjectContext использует
            // IncludeDependencyAnalyzerService, который в свою очередь подписан на OnProjectLoaded,
            // т.е. нам сначала необходимо дождаться пока обновиться IncludeDependencyAnalyzerService
            // для перезагруженного проекта, и только потом выполнять рутину по переключению контекста.
            VsixThreadHelper.RunOnUiThread(() => {
                var associatedFreshSharedItem = freshSharedItemEntries
                    .Select(sharedItemEntry => sharedItemEntry.MultiState.As<SharedItem>())
                    .FirstOrDefault(sharedItem => string.Equals(
                        sharedItem.HierarchyItemEntry.BaseViewModel.FilePath,
                        _invalidatedPreviousDocumentFilePath,
                        StringComparison.OrdinalIgnoreCase
                        ));

                if (associatedFreshSharedItem != null) {
                    associatedFreshSharedItem.OpenWithProjectContext();
                }
            });
        }
    }
}
