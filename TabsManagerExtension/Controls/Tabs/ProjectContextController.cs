using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.ObjectModel;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;

using Helpers.Ex;
using TabsManagerExtension.State.Document;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Строит меню project context и переоткрывает документы в выбранном проекте.</summary>
    internal sealed class ProjectContextController {
        private readonly EnvDTE80.DTE2 _dte;
        private readonly VirtualMenuControl _virtualMenuControl;
        private readonly TabCollectionManager _tabCollectionManager;
        private readonly Helpers.Collections.GroupsSelectionCoordinator<TabItemsGroupBase, TabItemBase> _selectionCoordinator;

        public ProjectContextController(
            EnvDTE80.DTE2 dte,
            VirtualMenuControl virtualMenuControl,
            TabCollectionManager tabCollectionManager,
            Helpers.Collections.GroupsSelectionCoordinator<TabItemsGroupBase, TabItemBase> selectionCoordinator
            ) {
            _dte = dte;
            _virtualMenuControl = virtualMenuControl;
            _tabCollectionManager = tabCollectionManager;
            _selectionCoordinator = selectionCoordinator;
        }

        public void MoveToRelatedProject(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _virtualMenuControl.HideImmediately();

            if (parameter is DocumentProjectReferencesInfo.RefEntry reference) {
                this.MoveDocumentToProjectGroup(reference.DocumentEntryBase);
                return;
            }
            if (parameter is not DocumentProjectReferencesInfo.GroupContextEntry groupContext || !groupContext.CanSwitch) {
                return;
            }

            string? activeDocumentPath = _dte.ActiveDocument?.FullName;
            var contextSwitchPlan = this.BuildGroupContextSwitchPlan(groupContext.DocumentReferences);
            var activatedSources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var documentReference in groupContext.DocumentReferences.ToList()) {
                contextSwitchPlan.TryGetValue(documentReference, out var sourcePath);
                this.MoveDocumentToProjectGroup(
                    documentReference.DocumentEntryBase,
                    playFeedback: false,
                    preferredContextSwitchSourcePath: sourcePath,
                    activatedContextSwitchSources: activatedSources
                );
            }

            Console.Beep(frequency: 1000, duration: 300);
            VsixThreadHelper.RunOnUiThread(async () => {
                await Task.Delay(450);

                var activeDocument = _dte.Documents
                    .Cast<EnvDTE.Document>()
                    .FirstOrDefault(document => string.Equals(
                        document.FullName,
                        activeDocumentPath,
                        StringComparison.OrdinalIgnoreCase
                    ));

                if (activeDocument != null) {
                    activeDocument.Activate();
                }
                else if (!string.IsNullOrEmpty(activeDocumentPath)) {
                    Helpers.Diagnostic.Logger.LogDebug($"[ProjectContextController] Cannot restore active document '{activeDocumentPath}' because its current frame was not found.");
                }
            });
        }

        public void MoveToRelatedProjectFile(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _virtualMenuControl.HideImmediately();

            if (parameter is not ProjectContextSourceEntry sourceEntry) {
                return;
            }

            var projectEntry = sourceEntry.ProjectContext switch {
                DocumentProjectReferencesInfo.RefEntry reference => reference.ProjectEntry,
                DocumentProjectReferencesInfo.GroupContextEntry groupContext => groupContext.ProjectEntry,
                _ => null
            };
            if (projectEntry?.MultiState.Current is not VsShell.Project.LoadedProject loadedProject) {
                return;
            }

            var sourceDocument = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance
                .SourcesRepresentationsTable
                .GetDocumentByProjectAndDocumentPath(projectEntry.BaseViewModel, sourceEntry.SourcePath);
            if (sourceDocument?.BaseViewModel.HierarchyItemEntry.BaseViewModel is VsShell.Hierarchy.RealHierarchyItem hierarchyItem) {
                int result = VsShell.Utils.VsHierarchyUtils.ClickOnSolutionHierarchyItem(
                    loadedProject.ProjectHierarchy.VsRealHierarchy,
                    hierarchyItem.ItemId
                );

                ErrorHandler.ThrowOnFailure(result);
                return;
            }

            if (!File.Exists(sourceEntry.SourcePath)) {
                Helpers.Diagnostic.Logger.LogWarning($"[ProjectContextController] File '{sourceEntry.SourcePath}' does not exist.");
                return;
            }

            _dte.ItemOperations
                .OpenFile(sourceEntry.SourcePath, EnvDTE.Constants.vsViewKindTextView)
                .Activate();
        }

        public void ToggleIncludersMenu(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (parameter is not FrameworkElement anchor ||
                anchor.DataContext is not Helpers.MenuItemCommand menuItem ||
                menuItem.CommandParameterContext is not object projectContext) {

                return;
            }

            if (_virtualMenuControl.IsChildMenuOpen &&
                ReferenceEquals(_virtualMenuControl.CurrentChildMenuDataContext, projectContext)) {

                _virtualMenuControl.HideChild();
                return;
            }

            this.ShowIncludersMenu(anchor, projectContext);
        }

        public void HandleMenuItemMouseEnter(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!_virtualMenuControl.IsChildMenuOpen ||
                parameter is not MenuItem menuItem ||
                menuItem.DataContext is not Helpers.MenuItemCommand menuItemCommand) {

                return;
            }

            object projectContext = menuItemCommand.CommandParameterContext;
            if (!ReferenceEquals(_virtualMenuControl.CurrentChildMenuDataContext, projectContext)) {
                this.ShowIncludersMenu(menuItem, projectContext);
            }
        }

        public void AppendProjectMenuItems(
            List<Helpers.IMenuItem> menuItems,
            TabItemDocument anchorTabItem,
            TabMenuItemFactory menuItemFactory
            ) {
            if (this.TryGetSelectedHeaderGroup(anchorTabItem, out var selectedHeaders)) {
                var groupContexts = this.BuildGroupContextEntries(selectedHeaders);
                if (groupContexts.Count == 0) {
                    return;
                }

                menuItems.Add(new Helpers.MenuItemSeparator());
                var commonContexts = groupContexts.Where(context => context.IsAvailableForAll).ToList();
                var differingContexts = groupContexts.Where(context => !context.IsAvailableForAll).ToList();
                this.AppendContextItems(menuItems, commonContexts, menuItemFactory);
                if (commonContexts.Count > 0 && differingContexts.Count > 0) {
                    menuItems.Add(new Helpers.MenuItemSeparator());
                }

                this.AppendContextItems(menuItems, differingContexts, menuItemFactory);
                return;
            }

            var references = anchorTabItem.DocumentProjectReferencesInfo.GetAvailableReferences();
            if (references.Count == 0) {
                return;
            }

            menuItems.Add(new Helpers.MenuItemSeparator());
            foreach (var reference in references) {
                menuItems.Add(menuItemFactory.CreateProjectContextItem(
                    reference.ProjectEntry.BaseViewModel.Name,
                    reference
                ));
            }

            menuItems.Add(menuItemFactory.CreateReloadProjectsItem(anchorTabItem.DocumentProjectReferencesInfo));
        }

        public void ReloadProjects(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (parameter is not DocumentProjectReferencesInfo referencesInfo) {
                return;
            }

            foreach (var reference in referencesInfo.References) {
                var projectViewModel = reference.ProjectEntry.BaseViewModel;
                if (projectViewModel is VsShell.Project.UnloadedProject) {
                    VsShell.Utils.VsHierarchyUtils.ReloadProject(projectViewModel.ProjectGuid);
                }
            }
        }

        private void MoveDocumentToProjectGroup(
            VsShell.Document.DocumentEntryBase documentEntry,
            bool playFeedback = true,
            string? preferredContextSwitchSourcePath = null,
            ISet<string>? activatedContextSwitchSources = null
            ) {
            ThreadHelper.ThrowIfNotOnUIThread();

            OpenWithProjectContext(
                documentEntry,
                playFeedback,
                preferredContextSwitchSourcePath,
                activatedContextSwitchSources
            );

            var documentViewModel = documentEntry.BaseViewModel;
            string filePath = documentViewModel.HierarchyItemEntry.BaseViewModel.FilePath;
            var tabItemDocument = _tabCollectionManager.Find(filePath);
            if (tabItemDocument == null) {
                return;
            }

            _tabCollectionManager.Remove(tabItemDocument);
            _tabCollectionManager.AddToGroup(
                tabItemDocument,
                new TabItemsDefaultGroup(documentViewModel.ProjectBaseViewModel.Name)
            );
        }

        private void ShowIncludersMenu(FrameworkElement anchor, object projectContext) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var sourcePaths = GetProjectContextSwitchSourcePaths(projectContext);
            string projectName = projectContext switch {
                DocumentProjectReferencesInfo.RefEntry reference => reference.ProjectEntry.BaseViewModel.Name,
                DocumentProjectReferencesInfo.GroupContextEntry groupContext => groupContext.ProjectEntry.BaseViewModel.Name,
                _ => "Project"
            };

            var childItems = new ObservableCollection<Helpers.IMenuItem> {
                new Helpers.MenuItemHeader { Header = projectName }
            };
            if (sourcePaths.Count == 0) {
                childItems.Add(new Helpers.MenuItemHeader { Header = "No transitive including files" });
            }
            else {
                foreach (string sourcePath in sourcePaths) {
                    childItems.Add(new Helpers.MenuItemCommand {
                        Header = Path.GetFileName(sourcePath),
                        Command = new Helpers.RelayCommand<object>(this.MoveToRelatedProjectFile),
                        CommandParameterContext = new ProjectContextSourceEntry(projectContext, sourcePath)
                    });
                }
            }

            Point screenPoint = anchor.ex_ToDpiAwareScreen(new Point(anchor.ActualWidth + 8, 0));
            _virtualMenuControl.ShowChild(screenPoint, projectContext, childItems);
        }

        private bool TryGetSelectedHeaderGroup(
            TabItemDocument anchorTabItem,
            out IReadOnlyList<TabItemDocument> selectedHeaders
            ) {
            var selectedItems = _selectionCoordinator.SelectedItems.Select(entry => entry.Item).ToList();
            bool anchorIsSelected = selectedItems.Any(item => ReferenceEquals(item, anchorTabItem));
            bool allSelectedItemsAreHeaders = selectedItems.All(
                item => item is TabItemDocument document && IsHeaderPath(document.FullName)
            );

            if (_selectionCoordinator.SelectionState != Helpers.Enums.SelectionState.Multiple ||
                !anchorIsSelected ||
                selectedItems.Count < 2 ||
                !allSelectedItemsAreHeaders) {

                selectedHeaders = Array.Empty<TabItemDocument>();
                return false;
            }

            selectedHeaders = selectedItems.Cast<TabItemDocument>().ToList();
            return true;
        }

        private IReadOnlyList<DocumentProjectReferencesInfo.GroupContextEntry> BuildGroupContextEntries(
            IReadOnlyList<TabItemDocument> selectedHeaders
            ) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var referencesByDocument = selectedHeaders
                .Select(header => header.DocumentProjectReferencesInfo
                    .GetAvailableReferences(includeSingleProject: true)
                    .GroupBy(reference => reference.ProjectEntry.BaseViewModel.ProjectGuid)
                    .ToDictionary(group => group.Key, group => group.First()))
                .ToList();
            var projectGuids = referencesByDocument.SelectMany(references => references.Keys).Distinct().ToList();
            var result = new List<DocumentProjectReferencesInfo.GroupContextEntry>();
            foreach (var projectGuid in projectGuids) {
                var documentReferences = referencesByDocument
                    .Where(references => references.ContainsKey(projectGuid))
                    .Select(references => references[projectGuid])
                    .ToList();

                result.Add(new DocumentProjectReferencesInfo.GroupContextEntry(
                    documentReferences[0].ProjectEntry,
                    documentReferences,
                    documentReferences.Count == selectedHeaders.Count
                ));
            }

            return result
                .OrderByDescending(context => context.IsAvailableForAll)
                .ThenBy(context => context.ProjectEntry.BaseViewModel.UniqueName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private IReadOnlyDictionary<DocumentProjectReferencesInfo.RefEntry, string> BuildGroupContextSwitchPlan(
            IReadOnlyList<DocumentProjectReferencesInfo.RefEntry> documentReferences
            ) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var candidatesByDocument = documentReferences.ToDictionary(
                reference => reference,
                reference => GetProjectContextSwitchSourcePaths(reference).ToHashSet(StringComparer.OrdinalIgnoreCase)
            );
            var openDocumentPaths = _dte.Documents
                .Cast<EnvDTE.Document>()
                .Select(document => document.FullName)
                .Where(path => !string.IsNullOrEmpty(path))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var uncoveredDocuments = candidatesByDocument
                .Where(pair => pair.Value.Count > 0)
                .Select(pair => pair.Key)
                .ToHashSet();
            var result = new Dictionary<DocumentProjectReferencesInfo.RefEntry, string>();

            while (uncoveredDocuments.Count > 0) {
                string selectedSourcePath = uncoveredDocuments
                    .SelectMany(reference => candidatesByDocument[reference])
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Select(path => new {
                        Path = path,
                        CoveredCount = uncoveredDocuments.Count(reference => candidatesByDocument[reference].Contains(path)),
                        IsOpen = openDocumentPaths.Contains(path)
                    })
                    .OrderByDescending(candidate => candidate.CoveredCount)
                    .ThenBy(candidate => candidate.IsOpen)
                    .ThenBy(candidate => candidate.Path, StringComparer.OrdinalIgnoreCase)
                    .First()
                    .Path;

                var coveredDocuments = uncoveredDocuments
                    .Where(reference => candidatesByDocument[reference].Contains(selectedSourcePath))
                    .ToList();
                foreach (var coveredDocument in coveredDocuments) {
                    result[coveredDocument] = selectedSourcePath;
                    uncoveredDocuments.Remove(coveredDocument);
                }
            }

            foreach (var reference in candidatesByDocument.Where(pair => pair.Value.Count == 0)) {
                Helpers.Diagnostic.Logger.LogWarning($"[ProjectContextController] No transitive .cpp was found for '{reference.Key.DocumentEntryBase.BaseViewModel.HierarchyItemEntry.BaseViewModel.FilePath}'.");
            }

            return result;
        }

        private static IReadOnlyList<string> GetProjectContextSwitchSourcePaths(object projectContext) {
            IEnumerable<DocumentProjectReferencesInfo.RefEntry> references = projectContext switch {
                DocumentProjectReferencesInfo.RefEntry reference => new[] { reference },
                DocumentProjectReferencesInfo.GroupContextEntry groupContext => groupContext.DocumentReferences,
                _ => Array.Empty<DocumentProjectReferencesInfo.RefEntry>()
            };

            return references
                .SelectMany(GetProjectContextSwitchSourcePaths)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
                .ThenBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static IReadOnlyList<string> GetProjectContextSwitchSourcePaths(
            DocumentProjectReferencesInfo.RefEntry reference
            ) {
            var document = GetCurrentProjectContextDocument(reference.DocumentEntryBase);
            var targetProject = reference.ProjectEntry.MultiState.Current as VsShell.Project.LoadedProject;
            return document != null && targetProject != null
                ? document.GetProjectContextSwitchSourcePaths(targetProject)
                : Array.Empty<string>();
        }

        private static VsShell.Document.Document? GetCurrentProjectContextDocument(
            VsShell.Document.DocumentEntryBase entry
            ) {
            return entry switch {
                VsShell.Document.ExternalIncludeEntry externalIncludeEntry =>
                    externalIncludeEntry.MultiState.Current as VsShell.Document.Document,
                VsShell.Document.SharedItemEntry sharedItemEntry =>
                    sharedItemEntry.MultiState.Current as VsShell.Document.Document,
                _ => null
            };
        }

        private static bool IsHeaderPath(string path) {
            string extension = Path.GetExtension(path);
            return string.Equals(extension, ".h", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".hpp", StringComparison.OrdinalIgnoreCase);
        }

        private static void OpenWithProjectContext(
            VsShell.Document.DocumentEntryBase documentEntry,
            bool playFeedback,
            string? preferredContextSwitchSourcePath,
            ISet<string>? activatedContextSwitchSources
            ) {
            if (documentEntry is VsShell.Document.ExternalIncludeEntry externalIncludeEntry) {
                if (externalIncludeEntry.MultiState.Current is VsShell.Document.ExternalInclude externalInclude) {
                    externalInclude.OpenWithProjectContext(
                        preferredContextSwitchSourcePath,
                        activatedContextSwitchSources
                    );
                    if (playFeedback) {
                        Console.Beep(frequency: 1000, duration: 300);
                    }
                }
                else if (externalIncludeEntry.MultiState.Current is VsShell.Document.InvalidatedDocument) {
                    System.Diagnostics.Debugger.Break();
                }

                return;
            }

            if (documentEntry is VsShell.Document.SharedItemEntry sharedItemEntry) {
                if (sharedItemEntry.MultiState.Current is VsShell.Document.SharedItem sharedItem) {
                    sharedItem.OpenWithProjectContext(
                        preferredContextSwitchSourcePath,
                        activatedContextSwitchSources
                    );
                    if (playFeedback) {
                        Console.Beep(frequency: 1000, duration: 300);
                    }
                }
                else if (sharedItemEntry.MultiState.Current is VsShell.Document.InvalidatedDocument invalidatedDocument) {
                    invalidatedDocument.OpenWithProjectContext();
                }
            }
        }

        private void AppendContextItems(
            ICollection<Helpers.IMenuItem> menuItems,
            IEnumerable<DocumentProjectReferencesInfo.GroupContextEntry> contexts,
            TabMenuItemFactory menuItemFactory
            ) {
            foreach (var context in contexts) {
                menuItems.Add(menuItemFactory.CreateProjectContextItem(
                    context.ProjectEntry.BaseViewModel.Name,
                    context
                ));
            }
        }

        private sealed class ProjectContextSourceEntry {
            public object ProjectContext { get; }
            public string SourcePath { get; }

            public ProjectContextSourceEntry(object projectContext, string sourcePath) {
                this.ProjectContext = projectContext;
                this.SourcePath = sourcePath;
            }
        }
    }
}
