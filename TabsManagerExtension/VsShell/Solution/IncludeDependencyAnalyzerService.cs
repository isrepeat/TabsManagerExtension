using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.VCCodeModel;
using Microsoft.VisualStudio.VCProjectEngine;


namespace TabsManagerExtension.VsShell.Solution.Services {
    public class IncludeDependencyAnalyzerService :
        TabsManagerExtension.Services.SingletonServiceBase<IncludeDependencyAnalyzerService>,
        TabsManagerExtension.Services.IExtensionService {

        private MsBuildSolutionWatcher _msBuildSolutionWatcher = null!;
        private SolutionSourceFileGraph _solutionSourceFileGraph = null!;
        private Helpers.DirectoryWatcher? _solutionDirWatcher;
        private DispatcherTimer _delayedFileChangeTimer = null!;
        private DispatcherTimer _graphBuildQuietTimer = null!;
        private DispatcherTimer _graphBuildMaxWaitTimer = null!;

        private readonly HashSet<Helpers.DirectoryChangedEventArgs> _pendingChangedFiles = new();

        private string? _lastLoadedSolutionName;
        private string? _initialAnalysisCompletedSolutionName;
        private bool _solutionHierarchyReady;
        private bool _buildingSolutionGraphInProcess = false;
        private bool _buildingProjectGraphInProcess = false;
        private int _msBuildRefreshQueued = 0;
        private volatile bool _isShuttingDown = false;
        private CancellationTokenSource? _graphBuildCancellation;
        private readonly object _includeParseCacheLock = new();
        private readonly Dictionary<string, IncludeParseCacheEntry> _includeParseCache = new(StringComparer.OrdinalIgnoreCase);

        private const string GraphBuildOperationKey = "IncludeDependencyGraph";

        private static readonly HashSet<string> _supportedCppExtensions = new(StringComparer.OrdinalIgnoreCase) {
            ".h", ".hh", ".hpp", ".hxx", ".inl", ".inc", ".ipp", ".tpp",
            ".c", ".cc", ".cpp", ".cxx", ".c++", ".ixx", ".cppm",
            ".m", ".mm", ".cu", ".cuh",
        };

        public IncludeDependencyAnalyzerService() { }

        //
        // IExtensionService
        //
        public IReadOnlyList<Type> DependsOn() {
            return new[] {
                typeof(VsShell.Services.VsIDEStateFlagsTrackerService),
                typeof(VsShell.Solution.Services.VsSolutionEventsTrackerService),
                typeof(VsShell.Solution.Services.SolutionHierarchyAnalyzerService),
                typeof(TabsManagerExtension.Services.ExtensionStatusService),
            };
        }

        public void Initialize() {
            ThreadHelper.ThrowIfNotOnUIThread();

            // DispatcherTimer выполняет отложенную обработку изменений файлов в UI-потоке.
            // Создаём его до InvokeForLastHandlerIfTriggered(): этот вызов немедленно сообщает сервису
            // о ранее открытом решении и может запустить наблюдение за файлами прямо внутри Initialize().
            // Если таймер создать позже, первое событие изменения файла обратится к полю со значением null.
            _isShuttingDown = false;
            _delayedFileChangeTimer = new DispatcherTimer {
                Interval = TimeSpan.FromMilliseconds(300)
            };
            _delayedFileChangeTimer.Tick += (_, _) => this.OnDelayedFileChangeTimerTick();

            _graphBuildQuietTimer = new DispatcherTimer {
                Interval = TimeSpan.FromSeconds(2)
            };
            // Запускаем граф после короткого периода без изменений hierarchy. Предельный таймер
            // гарантирует старт даже для решений, которые продолжают генерировать фоновые события.
            _graphBuildQuietTimer.Tick += (_, _) => this.TryStartScheduledGraphBuild("hierarchy quiet period");

            _graphBuildMaxWaitTimer = new DispatcherTimer {
                Interval = TimeSpan.FromSeconds(15)
            };
            _graphBuildMaxWaitTimer.Tick += (_, _) => this.TryStartScheduledGraphBuild("maximum wait timeout");

            VsShell.Services.VsIDEStateFlagsTrackerService.Instance.SolutionLoaded.Add(this.OnSolutionLoaded);
            VsShell.Services.VsIDEStateFlagsTrackerService.Instance.SolutionLoaded.InvokeForLastHandlerIfTriggered();
            VsShell.Services.VsIDEStateFlagsTrackerService.Instance.SolutionClosed.Add(this.OnSolutionClosed);
            VsShell.Services.VsIDEStateFlagsTrackerService.Instance.SolutionClosed.InvokeForLastHandlerIfTriggered();

            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.ProjectLoaded += this.OnProjectLoaded;
            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.ProjectUnloaded += this.OnProjectUnloaded;
            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.BackgroundSolutionLoadCompleted += this.OnBackgroundSolutionLoadCompleted;
            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.SolutionHierarchyActivity += this.OnSolutionHierarchyActivity;
            VsShell.Project.ProjectHierarchyTracker.AnyHierarchyChanged += this.OnSolutionHierarchyActivity;
            // Событие запоминает завершённый первичный снимок и сразу вызывает позднего подписчика.
            VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance.InitialAnalysisCompleted.Add(this.OnInitialHierarchyAnalysisCompleted);
            VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance.InitialAnalysisCompleted.InvokeForLastHandlerIfTriggered();

            Helpers.Diagnostic.Logger.LogDebug("[IncludeDependencyAnalyzerService] Initialized.");
        }

        public void Shutdown() {
            ThreadHelper.ThrowIfNotOnUIThread();

            _isShuttingDown = true;
            this.ClearAnalysisState();

            VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance.InitialAnalysisCompleted.Remove(this.OnInitialHierarchyAnalysisCompleted);
            VsShell.Project.ProjectHierarchyTracker.AnyHierarchyChanged -= this.OnSolutionHierarchyActivity;
            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.SolutionHierarchyActivity -= this.OnSolutionHierarchyActivity;
            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.BackgroundSolutionLoadCompleted -= this.OnBackgroundSolutionLoadCompleted;
            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.ProjectUnloaded -= this.OnProjectUnloaded;
            VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.ProjectLoaded -= this.OnProjectLoaded;
            VsShell.Services.VsIDEStateFlagsTrackerService.Instance.SolutionClosed.Remove(this.OnSolutionClosed);
            VsShell.Services.VsIDEStateFlagsTrackerService.Instance.SolutionLoaded.Remove(this.OnSolutionLoaded);

            ClearInstance();
            Helpers.Diagnostic.Logger.LogDebug("[IncludeDependencyAnalyzerService] Disposed.");
        }


        //
        // ░ Api
        // ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
        public bool IsReady() {
            return _solutionSourceFileGraph != null && _buildingSolutionGraphInProcess == false && _buildingProjectGraphInProcess == false;
        }

        private bool CanStartGraphBuild() {
            return
                !_isShuttingDown &&
                !string.IsNullOrEmpty(_lastLoadedSolutionName) &&
                !_buildingSolutionGraphInProcess &&
                !_buildingProjectGraphInProcess &&
                _solutionSourceFileGraph == null;
        }

        private bool StartGraphBuild() {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!this.CanStartGraphBuild()) {
                return false;
            }

            _graphBuildQuietTimer.Stop();
            _graphBuildMaxWaitTimer.Stop();
            _buildingSolutionGraphInProcess = true;
            this.ReportGraphProgress("Preparing", 1);
            Helpers.Diagnostic.Logger.LogDebug("[IncludeDependencyAnalyzerService] Graph build requested.");
            VsixThreadHelper.RunOnVsThread(this.BuildSolutionGraphAsync);
            return true;
        }


        public IReadOnlyCollection<VsShell.Project.LoadedProject> GetTransitiveProjectsIncludersByIncludeString(string includeString) {
            var transitiveIncluders = this.GetTransitiveFilesIncludersByIncludeString(includeString);
            return transitiveIncluders
                .Select(sf => sf.LoadedProject)
                .Distinct()
                .ToList();
        }


        public IReadOnlyCollection<VsShell.Project.LoadedProject> GetTransitiveProjectsIncludersByIncludePath(string includePath) {
            var transitiveIncluders = this.GetTransitiveFilesIncludersByIncludePath(includePath);
            return transitiveIncluders
                .Select(sf => sf.LoadedProject)
                .Distinct()
                .ToList();
        }


        /// <summary>
        /// Находит все <see cref="SourceFile"/>'ы, которые транзитивно включают include с заданным именем.
        /// </summary>
        /// <remarks>
        /// Используется двухфазный алгоритм: сначала ищутся прямые попадания по <c>RawInclude</c>,
        /// затем выполняется транзитивный обход вверх по цепочке включений.
        ///
        /// Поиск намеренно неточный: если в solution существуют <c>Game/Logger.h</c> и
        /// <c>Helpers.Shared/Logger.h</c>, результат объединит зависимости обоих файлов.
        /// Для поиска зависимостей конкретного физического файла используйте
        /// <see cref="GetTransitiveFilesIncludersByIncludePath(string)"/>.
        /// </remarks>
        /// <param name="includeString">Имя include-файла (например, <c>"Logger.h"</c>).</param>
        /// <returns>Список всех файлов, которые напрямую или транзитивно включают данный include.</returns>
        public IReadOnlyList<Document.SourceFile> GetTransitiveFilesIncludersByIncludeString(string includeString) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!this.IsReady()) {
                return Array.Empty<Document.SourceFile>();
            }

            this.RefreshEvaluationContextsIfNeeded();

            var result = new HashSet<Document.SourceFile>();
            var queue = new Queue<Document.SourceFile>();

            // ① Перебираем все ResolvedIncludeEntry во всех исходных файлах
            foreach (var kvp in _solutionSourceFileGraph.GetAllResolvedIncludeEntries()) {
                var sourceFile = kvp.Key;
                var resolvedList = kvp.Value;

                foreach (var resolved in resolvedList) {
                    var raw = resolved.IncludeEntry.RawInclude;

                    // ② Быстрая фильтрация по имени: проверяем RawInclude (например, "Logger.h")
                    if (!Path.GetFileName(raw).Equals(includeString, StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    // ③ Добавляем непосредственного включателя вместе с его project context.
                    // Например, Shared/X.h [Game] и Shared/X.h [Editor] — это разные узлы,
                    // даже если у них один физический FilePath.
                    if (resolved.ResolvedPath is not null) {
                        if (result.Add(sourceFile)) {
                            queue.Enqueue(sourceFile);
                        }
                    }
                }
            }

            // ④ Обратный обход: кто включает те файлы, что мы уже нашли (транзитивно вверх)
            while (queue.Count > 0) {
                var current = queue.Dequeue();

                // Продолжаем подъём только в том project context, в котором была разрешена текущая ветка.
                // Иначе Shared/X.h [Game] мог бы привести к Editor.cpp, включающему Shared/X.h [Editor].
                foreach (var includer in _solutionSourceFileGraph.GetIncludersOfResolvedPath(current.FilePath, current.LoadedProject)) {
                    if (result.Add(includer)) {
                        queue.Enqueue(includer); // продолжаем подниматься вверх по графу
                    }
                }
            }

            return result.ToList();
        }


        /// <summary>
        /// Находит все <see cref="SourceFile"/>'ы, которые транзитивно включают файл с заданным <c>ResolvedPath</c>.
        /// </summary>
        /// <remarks>
        /// Метод начинает с прямых включений указанного файла (по точному <c>ResolvedPath</c>),
        /// после чего выполняет транзитивный обход вверх — находит все исходные файлы,
        /// которые включают его опосредованно.
        ///
        /// Это наиболее точный и надёжный способ анализа, исключающий неоднозначности,
        /// возникающие при использовании только имени include-файла.
        /// </remarks>
        /// <param name="includePath">
        /// Абсолютный путь до включаемого файла, например:
        /// <c>"d:\PROJECT\Helpers.Shared\Logger.h"</c>.
        /// </param>
        /// <returns>
        /// Список <see cref="SourceFile"/>-файлов, которые напрямую или транзитивно включают указанный путь.
        /// </returns>
        public IReadOnlyList<Document.SourceFile> GetTransitiveFilesIncludersByIncludePath(string includePath) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!this.IsReady()) {
                return Array.Empty<Document.SourceFile>();
            }

            this.RefreshEvaluationContextsIfNeeded();

            string normalizedIncludePath;
            try {
                normalizedIncludePath = Path.GetFullPath(includePath);
            }
            catch {
                return Array.Empty<Document.SourceFile>();
            }

            var directFiles = _solutionSourceFileGraph.GetIncludersOfResolvedPath(normalizedIncludePath).ToList();
            if (directFiles.Count == 0) {
                return Array.Empty<Document.SourceFile>();
            }

            var result = new HashSet<Document.SourceFile>(directFiles);
            var queue = new Queue<Document.SourceFile>(directFiles);

            while (queue.Count > 0) {
                var current = queue.Dequeue();

                // Начальные узлы собираются во всех проектах, но каждая дальнейшая ветка остаётся
                // в своём контексте. Пример: Shared/X.h [Game] не переходит в Editor.cpp через
                // физически тот же Shared/X.h [Editor].
                foreach (var includer in _solutionSourceFileGraph.GetIncludersOfResolvedPath(current.FilePath, current.LoadedProject)) {
                    if (result.Add(includer)) {
                        queue.Enqueue(includer);
                    }
                }
            }

            return result.ToList();
        }


        public void LogIncludeTree() {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!this.IsReady()) {
                return;
            }

            foreach (var sourceFile in _solutionSourceFileGraph.AllSourceFiles.OrderBy(f => f.FilePath)) {
                Helpers.Diagnostic.Logger.LogDebug($"[File] {sourceFile.FilePath}");

                var resolvedIncludes = _solutionSourceFileGraph.GetResolvedIncludes(sourceFile);
                if (resolvedIncludes.Any()) {
                    foreach (var resolvedInclude in resolvedIncludes) {
                        Helpers.Diagnostic.Logger.LogDebug($"  └─ #include {resolvedInclude.IncludeEntry}");
                    }
                }
                else {
                    Helpers.Diagnostic.Logger.LogDebug("  └─ (no includes)");
                }
            }
        }


        //
        // ░ Event handlers
        // ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
        private void OnSolutionLoaded(string solutionName) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnSolutionLoaded()");

            if (solutionName == _lastLoadedSolutionName && _solutionSourceFileGraph != null) {
                return;
            }
            _lastLoadedSolutionName = solutionName;
            _solutionHierarchyReady = false;
            // До публикации нового графа зависимые UI-функции считаются неготовыми.
            TabsManagerExtension.Services.ExtensionStatusService.Instance.SetFeatureReady(
                TabsManagerExtension.Services.ExtensionStatusService.IncludeGraphFeature,
                false
            );

            this.ReportGraphProgress("Waiting for projects to load", 0);
            Helpers.Diagnostic.Logger.LogDebug("[IncludeDependencyAnalyzerService] Waiting for initial solution hierarchy analysis.");

            if (string.Equals(
                _initialAnalysisCompletedSolutionName,
                solutionName,
                StringComparison.OrdinalIgnoreCase)) {

                _solutionHierarchyReady = true;
                this.ScheduleGraphBuildAfterHierarchySettles();
            }
        }


        private void OnSolutionClosed(string solutionName) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnSolutionClosed()");
            ThreadHelper.ThrowIfNotOnUIThread();

            // Сбрасываем имя вместе с графом. Без этого повторное открытие того же .sln считалось
            // дубликатом SolutionLoaded и оставляло SourceFile со ссылками на уже disposed-проекты.
            _lastLoadedSolutionName = null;
            _initialAnalysisCompletedSolutionName = null;
            _solutionHierarchyReady = false;
            this.ClearAnalysisState();
            TabsManagerExtension.Services.ExtensionStatusService.Instance.SetFeatureReady(
                TabsManagerExtension.Services.ExtensionStatusService.IncludeGraphFeature,
                false
            );

            TabsManagerExtension.Services.ExtensionStatusService.Instance.RemoveOperation(GraphBuildOperationKey);
        }

        private void OnBackgroundSolutionLoadCompleted() {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrEmpty(_lastLoadedSolutionName) || _solutionSourceFileGraph != null) {
                return;
            }

            if (!_solutionHierarchyReady) {
                return;
            }

            this.ScheduleGraphBuildAfterHierarchySettles();
        }

        private void OnInitialHierarchyAnalysisCompleted(string solutionName) {
            ThreadHelper.ThrowIfNotOnUIThread();

            _initialAnalysisCompletedSolutionName = solutionName;

            if (string.IsNullOrEmpty(_lastLoadedSolutionName) ||
                !string.Equals(solutionName, _lastLoadedSolutionName, StringComparison.OrdinalIgnoreCase) ||
                _solutionSourceFileGraph != null) {

                return;
            }

            _solutionHierarchyReady = true;
            this.ScheduleGraphBuildAfterHierarchySettles();
        }

        private void ScheduleGraphBuildAfterHierarchySettles() {
            // Любая последующая активность hierarchy перезапускает quiet timer, но не max wait timer.
            this.ReportGraphProgress("Waiting for project hierarchy to settle", 0);
            this.RestartGraphBuildQuietTimer();
            _graphBuildMaxWaitTimer.Stop();
            _graphBuildMaxWaitTimer.Start();
        }

        private void OnSolutionHierarchyActivity() {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_solutionHierarchyReady && this.CanStartGraphBuild()) {
                this.RestartGraphBuildQuietTimer();
            }
        }

        private void RestartGraphBuildQuietTimer() {
            _graphBuildQuietTimer.Stop();
            _graphBuildQuietTimer.Start();
        }

        private void TryStartScheduledGraphBuild(string reason) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _graphBuildQuietTimer.Stop();
            _graphBuildMaxWaitTimer.Stop();

            if (!this.CanStartGraphBuild()) {
                return;
            }

            Helpers.Diagnostic.Logger.LogDebug($"[IncludeDependencyAnalyzerService] Starting graph build after {reason}.");
            this.StartGraphBuild();
        }


        private void OnProjectLoaded(_EventArgs.ProjectHierarchyChangedEventArgs e) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnProjectLoaded()");
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_solutionSourceFileGraph == null) {
                return;
            }

            var projectRealHierarchy = e.NewHierarchy.MultiState.As<Hierarchy.RealHierarchyItem>();
            PackageServices.VsSolution.GetGuidOfProject(projectRealHierarchy.VsRealHierarchy, out var projectGuid);

            var solutionHierarchyAnalyzer = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance;
            var loadedProject = solutionHierarchyAnalyzer.LoadedProjects
                .FirstOrDefault(p => p.ProjectGuid == projectGuid);

            if (loadedProject == null) {
                Helpers.Diagnostic.Logger.LogWarning($"Skip project graph update: project '{projectGuid}' is absent from the solution hierarchy analyzer.");
                return;
            }

            _msBuildSolutionWatcher?.AddOrUpdateProject(loadedProject.ShellProject.dteProject);
            this.UpdateProjectGraph(loadedProject);
        }


        private void OnProjectUnloaded(_EventArgs.ProjectHierarchyChangedEventArgs e) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnProjectUnloaded()");
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_solutionSourceFileGraph == null) {
                return;
            }

            var projectRealHierarchy = e.OldHierarchy.MultiState.As<Hierarchy.RealHierarchyItem>();
            PackageServices.VsSolution.GetGuidOfProject(projectRealHierarchy.VsRealHierarchy, out var projectGuid);

            var filesToRemove = _solutionSourceFileGraph.AllSourceFiles
                .Where(sf => sf.LoadedProject.ProjectGuid == projectGuid)
                .ToList();

            string? projectPath = filesToRemove.FirstOrDefault()?.LoadedProject.FullName;
            if (string.IsNullOrEmpty(projectPath)) {
                projectPath = Utils.EnvDteUtils.GetDteProjectFromHierarchy(projectRealHierarchy.VsRealHierarchy)?.FullName;
            }

            foreach (var sf in filesToRemove) {
                _solutionSourceFileGraph.RemoveSourceFile(sf);
            }

            if (!string.IsNullOrEmpty(projectPath)) {
                _msBuildSolutionWatcher?.RemoveProject(projectPath!);
            }
        }


        private void OnMsBuildIncludeEnvironmentChanged(string projectPath) {
            if (_isShuttingDown || Interlocked.CompareExchange(ref _msBuildRefreshQueued, 1, 0) != 0) {
                return;
            }

            VsixThreadHelper.RunOnUiThread(
                _delayedFileChangeTimer.Dispatcher,
                new Action(
                    () => {
                        Interlocked.Exchange(ref _msBuildRefreshQueued, 0);
                        if (_isShuttingDown || _solutionSourceFileGraph == null) {
                            return;
                        }

                        // Перечитываем все проекты: изменение PublicIncludeDirectories в одном
                        // referenced-проекте способно изменить разрешение include у его потребителей.
                        this.RefreshAllProjectGraphs();
                    }
                ),
                DispatcherPriority.Background
            );
        }


        private void OnSolutionDirectoryChanged(Helpers.DirectoryChangedEventArgs e) {
            if (_isShuttingDown) {
                return;
            }

            lock (_pendingChangedFiles) {
                _pendingChangedFiles.Add(e);
            }

            // Событие об изменении файла приходит из фонового потока. DispatcherTimer можно безопасно
            // запускать только в UI-потоке, в котором он был создан. BeginInvoke передаёт перезапуск
            // таймера в очередь UI-потока; без этого два потока могли одновременно менять его состояние.
            VsixThreadHelper.RunOnUiThread(
                _delayedFileChangeTimer.Dispatcher,
                new Action(
                    () => {
                        if (_isShuttingDown) {
                            return;
                        }

                        _delayedFileChangeTimer.Stop();
                        _delayedFileChangeTimer.Start();
                    }
                ),
                DispatcherPriority.Background
            );
        }


        private void OnDelayedFileChangeTimerTick() {
            _delayedFileChangeTimer.Stop();

            if (_isShuttingDown) {
                return;
            }

            List<Helpers.DirectoryChangedEventArgs> changedFiles;
            lock (_pendingChangedFiles) {
                changedFiles = _pendingChangedFiles.ToList();
                _pendingChangedFiles.Clear();
            }

            try {
                this.ProcessChangedProjectInputs(ref changedFiles);

                bool includeResolutionInvalidated = false;
                foreach (var changedFile in changedFiles) {
                    includeResolutionInvalidated |= this.ProcessChangedFile(changedFile);
                }

                if (includeResolutionInvalidated) {
                    this.ReresolveAllSourceFiles();
                }
            }
            catch (Exception ex) {
                // FileSystemWatcher сообщает о промежуточных состояниях atomic save. Ошибка чтения
                // временно заблокированного файла не должна падать из UI DispatcherTimer.
                Helpers.Diagnostic.Logger.LogError($"[IncludeDependencyAnalyzerService] Failed to process file changes: {ex}");
            }
        }


        //
        // ░ Internal logic
        // ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
        //
        /// <summary>
        /// Depends on SolutionHierarchyAnalyzerService.
        /// </summary>
        private async Task BuildSolutionGraphAsync() {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("BuildSolutionGraph()");
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            this.ClearAnalysisState();
            _buildingSolutionGraphInProcess = true;
            var cancellation = new CancellationTokenSource();
            _graphBuildCancellation = cancellation;
            CancellationToken cancellationToken = cancellation.Token;
            MsBuildSolutionWatcher? unpublishedWatcher = null;

            try {
                var solutionHierarchyAnalyzer = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance;
                var loadedProjectNodes = solutionHierarchyAnalyzer.LoadedProjects
                    .Where(this.IsCppProjectNode)
                    .ToList();

                var dteProjects = loadedProjectNodes
                    .Select(pn => pn.ShellProject.dteProject)
                    .ToList();

                var projectEvaluationInputs = MsBuildSolutionWatcher.CaptureProjectEvaluationInputs(dteProjects);

                this.ReportGraphProgress("Collecting project files", 3);
                var snapshotWatch = Stopwatch.StartNew();
                var projectSnapshots = await this.CreateProjectSnapshotsAsync(loadedProjectNodes, cancellationToken);
                snapshotWatch.Stop();

                var graphWatch = Stopwatch.StartNew();
                this.ReportGraphProgress("Building include graph", 25);
                // Парсинг файлов и построение рёбер не обращаются к DTE и выполняются вне UI-потока.
                var buildResult = await Task.Run(
                    () => {
                        var watcher = new MsBuildSolutionWatcher(projectEvaluationInputs);
                        unpublishedWatcher = watcher;
                        var newGraph = this.BuildGraphFromSnapshots(projectSnapshots, watcher, cancellationToken);
                        return (Watcher: watcher, Graph: newGraph);
                    },
                    cancellationToken
                );

                graphWatch.Stop();
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                buildResult.Watcher.AttachDteProjects(projectEvaluationInputs);
                buildResult.Watcher.IncludeEnvironmentChanged += this.OnMsBuildIncludeEnvironmentChanged;
                this.ReportGraphProgress("Publishing include graph", 98);
                // Готовый watcher и граф публикуются вместе; только после этого открываем UI-функции.
                _msBuildSolutionWatcher = buildResult.Watcher;
                _solutionSourceFileGraph = buildResult.Graph;
                unpublishedWatcher = null;

                string? solutionDir = Path.GetDirectoryName(PackageServices.Dte2.Solution.FullName);
                if (solutionDir != null && Directory.Exists(solutionDir)) {
                    _solutionDirWatcher = new Helpers.DirectoryWatcher(solutionDir);
                    _solutionDirWatcher.DirectoryChanged += this.OnSolutionDirectoryChanged;
                }

                int sourceFileCount = projectSnapshots.Sum(snapshot => snapshot.FilePaths.Count);
                Helpers.Diagnostic.Logger.LogDebug(
                    $"[IncludeDependencyAnalyzerService] Graph ready: projects={projectSnapshots.Count}, " +
                    $"source representations={sourceFileCount}, snapshot={snapshotWatch.ElapsedMilliseconds} ms, " +
                    $"background build={graphWatch.ElapsedMilliseconds} ms."
                );
                TabsManagerExtension.Services.ExtensionStatusService.Instance.SetFeatureReady(
                    TabsManagerExtension.Services.ExtensionStatusService.IncludeGraphFeature,
                    true
                );

                this.ReportGraphProgress("Ready", 100);
                TabsManagerExtension.Services.ExtensionStatusService.Instance.RemoveOperation(GraphBuildOperationKey);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                unpublishedWatcher?.Dispose();
                Helpers.Diagnostic.Logger.LogDebug("[IncludeDependencyAnalyzerService] Background graph build canceled.");
                if (!_isShuttingDown) {
                    TabsManagerExtension.Services.ExtensionStatusService.Instance.RemoveOperation(GraphBuildOperationKey);
                }
            }
            catch (Exception ex) {
                unpublishedWatcher?.Dispose();
                Helpers.Diagnostic.Logger.LogError($"[BuildSolutionGraph] exception: {ex}");
                this.ClearAnalysisState();
                this.ReportGraphProgress("Failed — see Output for details", 100);
            }
            finally {
                if (ReferenceEquals(_graphBuildCancellation, cancellation)) {
                    _graphBuildCancellation.Dispose();
                    _graphBuildCancellation = null;
                    _buildingSolutionGraphInProcess = false;
                }
            }
        }

        private async Task<List<ProjectFilesSnapshot>> CreateProjectSnapshotsAsync(
            IReadOnlyList<VsShell.Project.LoadedProject> loadedProjects,
            CancellationToken cancellationToken) {

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            var result = new List<ProjectFilesSnapshot>(loadedProjects.Count);

            for (int projectIndex = 0; projectIndex < loadedProjects.Count; projectIndex++) {
                var loadedProject = loadedProjects[projectIndex];
                cancellationToken.ThrowIfCancellationRequested();
                var filePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var stack = new Stack<EnvDTE.ProjectItem>();

                foreach (EnvDTE.ProjectItem item in loadedProject.ShellProject.dteProject.ProjectItems) {
                    stack.Push(item);
                }

                int processedItems = 0;
                while (stack.Count > 0) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var current = stack.Pop();

                    try {
                        if (current.FileCount > 0) {
                            string filePath = current.FileNames[1];
                            if (this.IsCppSourcePath(filePath) && File.Exists(filePath)) {
                                filePaths.Add(filePath);
                            }
                        }

                        if (current.ProjectItems != null) {
                            foreach (EnvDTE.ProjectItem child in current.ProjectItems) {
                                stack.Push(child);
                            }
                        }
                    }
                    catch (Exception ex) {
                        Helpers.Diagnostic.Logger.LogWarning($"[IncludeDependencyAnalyzerService] Skip unavailable project item: {ex.Message}");
                    }

                    processedItems++;
                    if (processedItems % 200 == 0) {
                        await Task.Yield();
                        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
                    }
                }

                result.Add(new ProjectFilesSnapshot(loadedProject, filePaths.ToList()));
                double progress = loadedProjects.Count == 0
                    ? 25
                    : 3 + ((projectIndex + 1) * 22.0 / loadedProjects.Count);
                this.ReportGraphProgress($"Collecting project files ({projectIndex + 1}/{loadedProjects.Count})", progress);
                await Task.Yield();
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            }

            return result;
        }

        private SolutionSourceFileGraph BuildGraphFromSnapshots(
            IReadOnlyList<ProjectFilesSnapshot> projectSnapshots,
            MsBuildSolutionWatcher msBuildSolutionWatcher,
            CancellationToken cancellationToken) {

            var graph = new SolutionSourceFileGraph(msBuildSolutionWatcher);
            int totalFiles = projectSnapshots.Sum(snapshot => snapshot.FilePaths.Count);
            int processedFiles = 0;
            int progressReportInterval = Math.Max(25, totalFiles / 100);
            foreach (var projectSnapshot in projectSnapshots) {
                foreach (string filePath in projectSnapshot.FilePaths) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var sourceFile = new Document.SourceFile(filePath, projectSnapshot.LoadedProject);
                    graph.AddSourceFileWithIncludes(sourceFile, this.ExtractRawIncludes(filePath));
                    processedFiles++;
                    if (processedFiles % progressReportInterval == 0 || processedFiles == totalFiles) {
                        double progress = totalFiles == 0 ? 95 : 25 + (processedFiles * 70.0 / totalFiles);
                        this.ReportGraphProgress($"Building include graph ({processedFiles}/{totalFiles})", progress);
                    }
                }
            }

            return graph;
        }


        private void UpdateProjectGraph(VsShell.Project.LoadedProject loadedProject) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("UpdateProjectGraph()");
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_solutionSourceFileGraph == null || !this.IsCppProjectNode(loadedProject)) {
                return;
            }

            _buildingProjectGraphInProcess = true;

            try {
                var discoveredProjectFiles = new HashSet<Document.SourceFile>();
                var stack = new Stack<EnvDTE.ProjectItem>();
                int skippedFileCount = 0;

                foreach (EnvDTE.ProjectItem item in loadedProject.ShellProject.dteProject.ProjectItems) {
                    stack.Push(item);
                }

                while (stack.Count > 0) {
                    var current = stack.Pop();

                    if (current.FileCount > 0) {
                        string filePath = current.FileNames[1];
                        string ext = Path.GetExtension(filePath);

                        bool isCppProjectFile = _supportedCppExtensions.Contains(ext);

                        if (isCppProjectFile) {
                            var newSourceFile = new Document.SourceFile(filePath, loadedProject);
                            var newIncludeEntries = this.ExtractRawIncludes(filePath);
                            discoveredProjectFiles.Add(newSourceFile);

                            if (_solutionSourceFileGraph.TryGetSourceFileRepresentations(filePath, out var candidates) &&
                                candidates.Any(sf => sf.LoadedProject.ProjectGuid == loadedProject.ProjectGuid)) {

                                _solutionSourceFileGraph.UpdateSourceFileWithIncludes(newSourceFile, newIncludeEntries);
                            }
                            else {
                                _solutionSourceFileGraph.AddSourceFileWithIncludes(newSourceFile, newIncludeEntries);
                            }
                        }
                        else {
                            skippedFileCount++;
                        }
                    }

                    // Всегда обходим вложенные элементы, даже если это была папка или файл без .h/.cpp
                    if (current.ProjectItems != null) {
                        foreach (EnvDTE.ProjectItem child in current.ProjectItems) {
                            stack.Push(child);
                        }
                    }
                }

                // UpdateProjectGraph должен быть синхронизацией, а не только add/update.
                // Например, если Old.cpp удалили из проекта, но оставили на диске, DTE больше не
                // обнаружит его, поэтому старое представление необходимо явно убрать из графа.
                var staleProjectFiles = _solutionSourceFileGraph.GetSourceFilesForProject(loadedProject.ProjectGuid)
                    .Where(sourceFile => !discoveredProjectFiles.Contains(sourceFile))
                    .ToList();

                foreach (var staleProjectFile in staleProjectFiles) {
                    _solutionSourceFileGraph.RemoveSourceFile(staleProjectFile);
                }

                Helpers.Diagnostic.Logger.LogDebug(
                    $"[IncludeDependencyAnalyzerService] Project graph synchronized: '{loadedProject.UniqueName}', " +
                    $"C/C++ files={discoveredProjectFiles.Count}, skipped items={skippedFileCount}, removed={staleProjectFiles.Count}."
                );
            }
            finally {
                _buildingProjectGraphInProcess = false;
            }
        }


        private List<Document.IncludeEntry> ExtractRawIncludes(string filePath) {
            var fileInfo = new FileInfo(filePath);
            long lastWriteTimeUtcTicks = fileInfo.LastWriteTimeUtc.Ticks;
            long length = fileInfo.Length;

            lock (_includeParseCacheLock) {
                if (_includeParseCache.TryGetValue(filePath, out var cached) &&
                    cached.LastWriteTimeUtcTicks == lastWriteTimeUtcTicks &&
                    cached.Length == length) {

                    return cached.IncludeEntries;
                }
            }

            // Один физический shared-файл может иметь представления в нескольких проектах.
            // Текст парсим один раз, а разрешение include всё равно выполняется в каждом project context.
            var includeEntries = CppIncludeParser.ParseFile(filePath);
            lock (_includeParseCacheLock) {
                _includeParseCache[filePath] = new IncludeParseCacheEntry(lastWriteTimeUtcTicks, length, includeEntries);
            }

            return includeEntries;
        }


        private void ProcessChangedProjectInputs(ref List<Helpers.DirectoryChangedEventArgs> changedFiles) {
            var changedProjectInputs = changedFiles
                .Where(changedFile => this.IsProjectInputPath(changedFile.FullPath))
                .ToList();

            if (changedProjectInputs.Count == 0) {
                return;
            }

            changedFiles.RemoveAll(changedProjectInputs.Contains);

            // .vcxitems и импортированные .props/.targets могут влиять сразу на несколько проектов.
            // Например, Helpers.Shared.vcxitems импортируется Game, Editor и Engine, поэтому обновления
            // только проекта с именем Helpers.Shared недостаточно — пересчитываем все loaded contexts.
            this.RefreshAllProjectGraphs();
        }


        private bool ProcessChangedFile(Helpers.DirectoryChangedEventArgs changedFile) {
            ThreadHelper.ThrowIfNotOnUIThread();

            bool isCurrentPathSupported = this.IsCppSourcePath(changedFile.FullPath);
            bool isOldPathSupported = changedFile.OldPath != null && this.IsCppSourcePath(changedFile.OldPath);
            if (!isCurrentPathSupported && !isOldPathSupported) {
                return false;
            }

            bool resolutionInvalidated = changedFile.ChangeType != Helpers.DirectoryChangeType.Changed;

            if (isOldPathSupported && changedFile.ChangeType == Helpers.DirectoryChangeType.Renamed) {
                this.RemoveSourceFileRepresentations(changedFile.OldPath!);
            }

            if (!isCurrentPathSupported) {
                return resolutionInvalidated;
            }

            if (!File.Exists(changedFile.FullPath)) {
                this.RemoveSourceFileRepresentations(changedFile.FullPath);
                return resolutionInvalidated;
            }

            if (changedFile.ChangeType == Helpers.DirectoryChangeType.Changed &&
                _solutionSourceFileGraph.TryGetSourceFileRepresentations(changedFile.FullPath, out var candidates) &&
                candidates.Count > 0) {

                this.UpdateIncludesIfNeeded(changedFile.FullPath, candidates);
            }
            else {
                this.SynchronizeSourceFileRepresentations(changedFile.FullPath);
            }

            return resolutionInvalidated;
        }


        private void SynchronizeSourceFileRepresentations(string filePath) {
            var solutionHierarchyAnalyzer = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance;
            var ownerProjects = solutionHierarchyAnalyzer.LoadedProjects
                .Where(loadedProject => Utils.EnvDteUtils.IsFileInProject(filePath, loadedProject.ShellProject.dteProject))
                .ToList();

            _solutionSourceFileGraph.TryGetSourceFileRepresentations(filePath, out var existingRepresentations);
            var ownerProjectGuids = ownerProjects
                .Select(project => project.ProjectGuid)
                .ToHashSet();

            foreach (var staleRepresentation in existingRepresentations.Where(sourceFile => !ownerProjectGuids.Contains(sourceFile.LoadedProject.ProjectGuid)).ToList()) {
                _solutionSourceFileGraph.RemoveSourceFile(staleRepresentation);
            }

            if (ownerProjects.Count == 0) {
                return;
            }

            var newIncludes = this.ExtractRawIncludes(filePath);
            foreach (var ownerProject in ownerProjects) {
                var sourceFile = new Document.SourceFile(filePath, ownerProject);
                if (existingRepresentations.Any(existing => existing.LoadedProject.ProjectGuid == ownerProject.ProjectGuid)) {
                    _solutionSourceFileGraph.UpdateSourceFileWithIncludes(sourceFile, newIncludes);
                }
                else {
                    _solutionSourceFileGraph.AddSourceFileWithIncludes(sourceFile, newIncludes);
                }

                Helpers.Diagnostic.Logger.LogDebug($"[synchronized] {filePath} [{ownerProject.UniqueName}] → graph updated by fswatcher");
            }
        }


        private void RemoveSourceFileRepresentations(string filePath) {
            if (!_solutionSourceFileGraph.TryGetSourceFileRepresentations(filePath, out var candidates)) {
                return;
            }

            foreach (var sourceFile in candidates.ToList()) {
                _solutionSourceFileGraph.RemoveSourceFile(sourceFile);
                Helpers.Diagnostic.Logger.LogDebug($"[removed] {filePath} [{sourceFile.LoadedProject.UniqueName}] → removed from graph");
            }
        }


        private void RefreshAllProjectGraphs() {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_solutionSourceFileGraph == null || _msBuildSolutionWatcher == null) {
                return;
            }

            var loadedProjects = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance.LoadedProjects.ToList();
            _msBuildSolutionWatcher.SynchronizeProjects(loadedProjects.Select(project => project.ShellProject.dteProject));

            foreach (var loadedProject in loadedProjects) {
                this.UpdateProjectGraph(loadedProject);
            }
        }


        private void ReresolveAllSourceFiles() {
            if (_solutionSourceFileGraph == null) {
                return;
            }

            // Создание, удаление или rename target-файла меняет результат разрешения include даже
            // тогда, когда текст включателя не изменился. Например, после появления Missing.h запись
            // "Missing.h → null" должна стать "Missing.h → C:/Project/Missing.h".
            var sourceFilesWithRawIncludes = _solutionSourceFileGraph.AllSourceFiles
                .Select(
                    sourceFile => (
                        SourceFile: sourceFile,
                        RawIncludes: _solutionSourceFileGraph.GetRawIncludes(sourceFile).ToList()
                    )
                )
                .ToList();

            foreach (var sourceFileWithIncludes in sourceFilesWithRawIncludes) {
                _solutionSourceFileGraph.UpdateSourceFileWithIncludes(
                    sourceFileWithIncludes.SourceFile,
                    sourceFileWithIncludes.RawIncludes
                );
            }
        }


        private void RefreshEvaluationContextsIfNeeded() {
            if (_msBuildSolutionWatcher != null && _msBuildSolutionWatcher.RefreshEvaluationContexts()) {
                // Переключение Debug|x64 → Release|x64 не меняет текст исходников, но может
                // полностью изменить AdditionalIncludeDirectories. Перед запросом обновляем edges,
                // чтобы результат соответствовал активной конфигурации Visual Studio.
                this.ReresolveAllSourceFiles();
            }
        }


        private bool IsCppSourcePath(string filePath) {
            return _supportedCppExtensions.Contains(Path.GetExtension(filePath));
        }

        private bool IsCppProjectNode(VsShell.Project.LoadedProject loadedProject) {
            string extension = Path.GetExtension(loadedProject.UniqueName);
            return string.Equals(extension, ".vcxproj", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".vcxitems", StringComparison.OrdinalIgnoreCase);
        }


        private bool IsProjectInputPath(string filePath) {
            string extension = Path.GetExtension(filePath);
            return string.Equals(extension, ".vcxproj", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".vcxitems", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".filters", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".props", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".targets", StringComparison.OrdinalIgnoreCase);
        }


        private void ClearAnalysisState() {
            _delayedFileChangeTimer?.Stop();
            _graphBuildQuietTimer?.Stop();
            _graphBuildMaxWaitTimer?.Stop();
            _graphBuildCancellation?.Cancel();

            // DirectoryWatcher и MSBuild watcher работают в фоновых потоках. Сначала отписываемся,
            // затем освобождаем ресурсы, чтобы закрытый solution не поставил новую UI-задачу в очередь.
            if (_solutionDirWatcher != null) {
                _solutionDirWatcher.DirectoryChanged -= this.OnSolutionDirectoryChanged;
                _solutionDirWatcher.Dispose();
                _solutionDirWatcher = null;
            }

            if (_msBuildSolutionWatcher != null) {
                _msBuildSolutionWatcher.IncludeEnvironmentChanged -= this.OnMsBuildIncludeEnvironmentChanged;
                _msBuildSolutionWatcher.Dispose();
                _msBuildSolutionWatcher = null!;
            }

            _solutionSourceFileGraph?.Clear();
            _solutionSourceFileGraph = null!;
            Interlocked.Exchange(ref _msBuildRefreshQueued, 0);
            _buildingProjectGraphInProcess = false;

            lock (_pendingChangedFiles) {
                _pendingChangedFiles.Clear();
            }
        }

        private void ReportGraphProgress(string stage, double progress) {
            TabsManagerExtension.Services.ExtensionStatusService.Instance.ReportOperation(
                GraphBuildOperationKey,
                "Dependency graph",
                stage,
                progress
            );
        }

        private sealed class ProjectFilesSnapshot {
            public VsShell.Project.LoadedProject LoadedProject { get; }
            public IReadOnlyList<string> FilePaths { get; }

            public ProjectFilesSnapshot(VsShell.Project.LoadedProject loadedProject, IReadOnlyList<string> filePaths) {
                this.LoadedProject = loadedProject;
                this.FilePaths = filePaths;
            }
        }

        private sealed class IncludeParseCacheEntry {
            public long LastWriteTimeUtcTicks { get; }
            public long Length { get; }
            public List<Document.IncludeEntry> IncludeEntries { get; }

            public IncludeParseCacheEntry(long lastWriteTimeUtcTicks, long length, List<Document.IncludeEntry> includeEntries) {
                this.LastWriteTimeUtcTicks = lastWriteTimeUtcTicks;
                this.Length = length;
                this.IncludeEntries = includeEntries;
            }
        }


        private void UpdateIncludesIfNeeded(string filePath, IReadOnlyList<Document.SourceFile> candidates) {
            var updated = new List<(Document.SourceFile OldFile, List<Document.IncludeEntry> NewIncludes)>();
            var newIncludeEntries = this.ExtractRawIncludes(filePath);

            foreach (var oldFile in candidates) {
                var oldIncludeEntries = _solutionSourceFileGraph.GetResolvedIncludes(oldFile)
                    .Select(resolvedInclude => resolvedInclude.IncludeEntry)
                    .ToList();

                bool changed = newIncludeEntries.Count != oldIncludeEntries.Count ||
                               !newIncludeEntries.SequenceEqual(oldIncludeEntries);

                if (changed) {
                    updated.Add((oldFile, newIncludeEntries));
                }
            }

            foreach (var (oldFile, changedIncludeEntries) in updated) {
                var updatedFile = new Document.SourceFile(filePath, oldFile.LoadedProject);
                _solutionSourceFileGraph.UpdateSourceFileWithIncludes(updatedFile, changedIncludeEntries);
                Helpers.Diagnostic.Logger.LogDebug($"[include changed] {filePath} [{oldFile.LoadedProject.UniqueName}] → includes updated");
            }
        }
    }
}
