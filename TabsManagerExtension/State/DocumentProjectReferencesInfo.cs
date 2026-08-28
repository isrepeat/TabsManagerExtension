using System;
using System.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using Microsoft.VisualStudio.Shell;

using Helpers.Attributes;

namespace TabsManagerExtension.State.Document {
    public partial class DocumentProjectReferencesInfo : Helpers.ObservableObject {
        private readonly ReadOnlyObservableCollection<DocumentProjectReferencesInfo.RefEntry> _readonlyReferences;
        public ReadOnlyObservableCollection<DocumentProjectReferencesInfo.RefEntry> References => _readonlyReferences;

        
        [ObservableProperty(AccessMarker.Get, AccessMarker.PrivateSet)]
        public bool _hasUnloadedProjects = false;

        
        private string _documenFullName;
        private ObservableCollection<DocumentProjectReferencesInfo.RefEntry> _references = new();
        private IReadOnlyList<DocumentProjectReferencesInfo.RefEntry> _availableReferencesIncludingSingleProject = Array.Empty<DocumentProjectReferencesInfo.RefEntry>();
        private Helpers.Time.DelayedEventsHandler _delayedEventsHandler;
        private bool _isInitialized;
        // Версия снимка таблиц SolutionHierarchyAnalyzerService, по которому был заполнен
        // _availableReferencesIncludingSingleProject. Значение -1 означает, что кэш ещё не строился.
        private int _representationsVersion = -1;

        public DocumentProjectReferencesInfo(string documenFullName) {
            _documenFullName = documenFullName;

            _references.CollectionChanged += this.OnReferencesCollectionChanged;
            _readonlyReferences = new ReadOnlyObservableCollection<DocumentProjectReferencesInfo.RefEntry>(_references);

            _delayedEventsHandler = new Helpers.Time.DelayedEventsHandler(TimeSpan.FromMilliseconds(300));
            _delayedEventsHandler.OnReady += this.OnRefreshReferences;
        }


        public void UpdateReferences() {
            foreach (var reference in _availableReferencesIncludingSingleProject) {
                reference.ProjectEntry.BaseViewModel.SharedItemsChanged.Remove(this.OnSharedItemsChanged);
                reference.ProjectEntry.BaseViewModel.ExternalIncludesChanged.Remove(this.OnExternalIncludesChanged);
            }
            _references.Clear();

            // Анализатор решения перестраивает свои таблицы по событиям решения и проектов.
            // Здесь только читаем готовые таблицы и обновляем локальный кэш документа.
            _availableReferencesIncludingSingleProject = this.BuildReferences(includeSingleProject: true);
            var refEntries = _availableReferencesIncludingSingleProject.Count < 2
                ? Array.Empty<DocumentProjectReferencesInfo.RefEntry>()
                : _availableReferencesIncludingSingleProject;

            foreach (var refEntry in refEntries) {
                _references.Add(refEntry);
            }

            // Подписываемся на изменения коллекций ExternalIncludes и SharedItems
            // поскольку они могут очищаться при выгрузке проектов и нам нужно заново
            // обновлять текущие элементы соотвутствующие новым значениям и этих коллекций.
            foreach (var reference in _availableReferencesIncludingSingleProject) {
                reference.ProjectEntry.BaseViewModel.ExternalIncludesChanged.Add(this.OnExternalIncludesChanged);
                reference.ProjectEntry.BaseViewModel.SharedItemsChanged.Add(this.OnSharedItemsChanged);
            }

            base.OnPropertyChanged(nameof(this.References));
            this.UpdateProperty(nameof(this.HasUnloadedProjects));
            // Фиксируем версию только после полного обновления обоих локальных списков.
            _representationsVersion = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance.RepresentationsVersion;
            _isInitialized = true;
        }


        public IReadOnlyList<DocumentProjectReferencesInfo.RefEntry> GetAvailableReferences(bool includeSingleProject = false) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var solutionHierarchyAnalyzer = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance;
            // Подписки проекта ловят только изменения уже известных references. Сравнение версий
            // покрывает добавление context-а, которого в первом (например, раннем) снимке не было.
            if (!_isInitialized || _representationsVersion != solutionHierarchyAnalyzer.RepresentationsVersion) {
                this.UpdateReferences();
            }

            return includeSingleProject
                ? _availableReferencesIncludingSingleProject
                : _readonlyReferences;
        }


        public void Clear() {
            foreach (var reference in _availableReferencesIncludingSingleProject) {
                reference.ProjectEntry.BaseViewModel.SharedItemsChanged.Remove(this.OnSharedItemsChanged);
                reference.ProjectEntry.BaseViewModel.ExternalIncludesChanged.Remove(this.OnExternalIncludesChanged);
            }

            _references.Clear();
            _availableReferencesIncludingSingleProject = Array.Empty<DocumentProjectReferencesInfo.RefEntry>();
            _representationsVersion = -1;
            _isInitialized = false;
        }


        private void OnReferencesCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e) {
            if (e.Action == NotifyCollectionChangedAction.Reset) {
                foreach (var oldItemRefEntry in _references) {
                    oldItemRefEntry.ProjectEntry.MultiState.StateChanged -= this.OnProjectMultiStateChanged;
                }
            }

            if (e.OldItems != null) {
                foreach (DocumentProjectReferencesInfo.RefEntry oldItemRefEntry in e.OldItems) {
                    oldItemRefEntry.ProjectEntry.MultiState.StateChanged -= this.OnProjectMultiStateChanged;
                }
            }
            if (e.NewItems != null) {
                foreach (DocumentProjectReferencesInfo.RefEntry newItemRefEntry in e.NewItems) {
                    newItemRefEntry.ProjectEntry.MultiState.StateChanged += this.OnProjectMultiStateChanged;
                }
            }
        }

        private void OnExternalIncludesChanged(IReadOnlyList<VsShell.Document.ExternalIncludeEntry> newExternalIncludes) {
            _delayedEventsHandler.Schedule();
        }

        private void OnSharedItemsChanged(IReadOnlyList<VsShell.Document.SharedItemEntry> newSharedItems) {
            _delayedEventsHandler.Schedule();
        }

        private void OnProjectMultiStateChanged() {
            _delayedEventsHandler.Schedule();
        }


        private void OnRefreshReferences() {
            this.UpdateReferences();
        }


        private void UpdateProperty(string? propertyName) {
            switch (propertyName) {
                case nameof(DocumentProjectReferencesInfo.HasUnloadedProjects):
                    this.HasUnloadedProjects = _references.Any(refEntry => !refEntry.ProjectEntry.IsLoaded);
                    break;
            }
        }


        private IReadOnlyList<DocumentProjectReferencesInfo.RefEntry> BuildReferences(bool includeSingleProject = false) {
            var ext = System.IO.Path.GetExtension(_documenFullName);
            switch (ext) {
                case ".h":
                case ".hpp":
                    break;

                default:
                    return Array.Empty<DocumentProjectReferencesInfo.RefEntry>();
            }

            var solutionHierarchyAnalyzer = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance;

            // Получаем все проекты, которые знают про этот файл.
            var externalIncludesSolutionProjectNodes = solutionHierarchyAnalyzer.ExternalIncludeRepresentationsTable
                .GetProjectsByDocumentPath(_documenFullName);

            var sharedItemsSolutionProjectNodes = solutionHierarchyAnalyzer.SharedItemsRepresentationsTable
                .GetProjectsByDocumentPath(_documenFullName);

            var allSolutionProjects = externalIncludesSolutionProjectNodes
                .Concat(sharedItemsSolutionProjectNodes)
                .ToList();


            if (!includeSingleProject && allSolutionProjects.Count < 2) {
                return Array.Empty<DocumentProjectReferencesInfo.RefEntry>(); // Игнорируем только лишь ссылки на собсвтенные проекты.
            }

            var refEntries = new List<DocumentProjectReferencesInfo.RefEntry>();

            //VsShell.Utils.VsHierarchyUtils.LogSolutionHierarchy();

            foreach (var projectEntry in allSolutionProjects) {
                var externalInclude = solutionHierarchyAnalyzer.ExternalIncludeRepresentationsTable
                    .GetDocumentByProjectAndDocumentPath(projectEntry, _documenFullName);

                if (externalInclude != null) {
                    refEntries.Add(new DocumentProjectReferencesInfo.RefEntry(projectEntry, externalInclude));
                }
            }


            foreach (var projectEntry in allSolutionProjects) {
                var sharedItemNode = solutionHierarchyAnalyzer.SharedItemsRepresentationsTable
                    .GetDocumentByProjectAndDocumentPath(projectEntry, _documenFullName);

                if (sharedItemNode != null) {
                    refEntries.Add(new DocumentProjectReferencesInfo.RefEntry(projectEntry, sharedItemNode));
                }
            }

            refEntries = refEntries
                .OrderBy(entry => entry.ProjectEntry.BaseViewModel.UniqueName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return refEntries;
        }


        // TODO: Add to CodeAnalyzer support generate nested partial classes.
        public class RefEntry : Helpers.ObservableObject {
            public VsShell.Project.ProjectEntry ProjectEntry { get; private set; }

            
            public VsShell.Document.DocumentEntryBase _documentEntryBase;
            public VsShell.Document.DocumentEntryBase DocumentEntryBase {
                get {
                    return _documentEntryBase;
                }
                private set {
                    if (_documentEntryBase != value) {
                        _documentEntryBase = value;
                    }
                    base.OnPropertyChanged();
                }
            }

            public RefEntry(
                VsShell.Project.ProjectEntry projectEntry,
                VsShell.Document.DocumentEntryBase documentEntryBase
                ) {
                Helpers.ThrowableAssert.Require(projectEntry.BaseViewModel.Equals(documentEntryBase.BaseViewModel.ProjectBaseViewModel));
                this.ProjectEntry = projectEntry;
                this.DocumentEntryBase = documentEntryBase;
            }

            public void UpdateDocumentEntry(VsShell.Document.DocumentEntryBase newDocumentEntryBase) {
                this.DocumentEntryBase = newDocumentEntryBase;
            }
        }


        public sealed class GroupContextEntry {
            public VsShell.Project.ProjectEntry ProjectEntry { get; }
            public IReadOnlyList<DocumentProjectReferencesInfo.RefEntry> DocumentReferences { get; }
            public bool IsAvailableForAll { get; }
            public bool CanSwitch { get; }

            public GroupContextEntry(
                VsShell.Project.ProjectEntry projectEntry,
                IReadOnlyList<DocumentProjectReferencesInfo.RefEntry> documentReferences,
                bool isAvailableForAll
                ) {
                this.ProjectEntry = projectEntry;
                this.DocumentReferences = documentReferences;
                this.IsAvailableForAll = isAvailableForAll;
                this.CanSwitch =
                    isAvailableForAll &&
                    projectEntry.IsLoaded &&
                    documentReferences.All(reference =>
                        !reference.DocumentEntryBase.IsInvalidated &&
                        !reference.DocumentEntryBase.BaseViewModel.IsDisposed);
            }
        }
    }
}
