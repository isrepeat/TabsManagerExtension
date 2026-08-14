using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Build.Locator;
using Microsoft.Build.Evaluation;
using Microsoft.VisualStudio.Shell;


namespace TabsManagerExtension.VsShell.Solution {
    public static class MsBuildEnvironment {
        private static bool _initialized = false;

        public static void EnsureInitialized() {
            if (_initialized) {
                return;
            }

            ThreadHelper.ThrowIfNotOnUIThread();

            bool isRunningInsideVisualStudio = PackageServices.TryGetDte2() is EnvDTE80.DTE2;
            if (!isRunningInsideVisualStudio) {
                /// Вызов MSBuildLocator.RegisterDefaults() необходим в обычных .NET приложениях,
                /// чтобы указать путь к используемой версии MSBuild. Однако в Visual Studio (внутри VSIX) MSBuild-сборки
                /// уже загружены, и попытка вызвать RegisterDefaults() вызовет исключение,
                /// т.к. Visual Studio уже настроена на нужную среду MSBuild.
                MSBuildLocator.RegisterDefaults();

                Environment.SetEnvironmentVariable(
                    "VCTargetsPath",
                    @"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Microsoft\VC\v170\"
                );
            }

            _initialized = true;
        }
    }


    public sealed class MsBuildProjectAnalyzer : IDisposable {
        public event Action<string>? IncludeEnvironmentChanged;

        private readonly object _stateLock = new();
        private readonly object _reloadLock = new();
        private readonly string _projectFilePath;
        private readonly FileSystemWatcher _projectWatcher;
        private readonly Dictionary<string, FileSystemWatcher> _dependencyDirectoryWatchers = new(StringComparer.OrdinalIgnoreCase);

        private Dictionary<string, string> _globalProperties;
        private List<string> _currentReferences = new();
        private List<string> _projectIncludeDirectories = new();
        private List<string> _publicIncludeDirsFromReferences = new();
        private HashSet<string> _watchedDependencyPaths = new(StringComparer.OrdinalIgnoreCase);
        private Microsoft.Build.Evaluation.Project? _loadedProject;
        private bool _disposed = false;

        public IReadOnlyList<string> IncludeDirectories {
            get {
                lock (_stateLock) {
                    return _projectIncludeDirectories
                        .Concat(_publicIncludeDirsFromReferences)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
            }
        }

        public MsBuildProjectAnalyzer(string projectFilePath, IReadOnlyDictionary<string, string> globalProperties) {
            MsBuildEnvironment.EnsureInitialized();

            _projectFilePath = Path.GetFullPath(projectFilePath);
            _globalProperties = globalProperties.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);

            _projectWatcher = new FileSystemWatcher(Path.GetDirectoryName(_projectFilePath)!) {
                Filter = Path.GetFileName(_projectFilePath),
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName,
                EnableRaisingEvents = true
            };

            _projectWatcher.Changed += this.OnWatchedProjectChanged;
            _projectWatcher.Created += this.OnWatchedProjectChanged;
            _projectWatcher.Renamed += this.OnWatchedProjectRenamed;
            _projectWatcher.Deleted += this.OnWatchedProjectChanged;

            this.ReloadProject(globalProperties);
        }

        public void Dispose() {
            lock (_reloadLock) {
                lock (_stateLock) {
                    if (_disposed) {
                        return;
                    }

                    _disposed = true;
                }

                _projectWatcher.Changed -= this.OnWatchedProjectChanged;
                _projectWatcher.Created -= this.OnWatchedProjectChanged;
                _projectWatcher.Renamed -= this.OnWatchedProjectRenamed;
                _projectWatcher.Deleted -= this.OnWatchedProjectChanged;
                _projectWatcher.Dispose();

                foreach (var dependencyWatcher in _dependencyDirectoryWatchers.Values) {
                    dependencyWatcher.Dispose();
                }
                _dependencyDirectoryWatchers.Clear();

                Microsoft.Build.Evaluation.Project? projectToUnload;
                lock (_stateLock) {
                    projectToUnload = _loadedProject;
                    _loadedProject = null;
                }

                MsBuildProjectAnalyzer.UnloadProject(projectToUnload);
            }
        }

        public void ReloadProject(IReadOnlyDictionary<string, string> globalProperties) {
            lock (_stateLock) {
                if (_disposed) {
                    return;
                }

                _globalProperties = globalProperties.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
            }

            this.ReloadProject();
        }

        public bool EnsureEvaluationContext(IReadOnlyDictionary<string, string> globalProperties) {
            bool contextChanged;
            lock (_stateLock) {
                contextChanged = !MsBuildProjectAnalyzer.EvaluationContextsEqual(_globalProperties, globalProperties);
            }

            if (contextChanged) {
                this.ReloadProject(globalProperties);
            }

            return contextChanged;
        }

        private void ReloadProject() {
            lock (_reloadLock) {
                Dictionary<string, string> globalProperties;
                lock (_stateLock) {
                    if (_disposed) {
                        return;
                    }

                    globalProperties = new Dictionary<string, string>(_globalProperties, StringComparer.OrdinalIgnoreCase);
                }

                ProjectCollection? newProjectCollection = null;
                Microsoft.Build.Evaluation.Project? newProject = null;

                try {
                    newProjectCollection = new ProjectCollection(globalProperties);
                    newProject = newProjectCollection.LoadProject(_projectFilePath);

                    var newReferences = newProject.GetItems("ProjectReference")
                        .Select(
                            item => Path.GetFullPath(
                                Path.Combine(
                                    Path.GetDirectoryName(_projectFilePath)!,
                                    item.EvaluatedInclude
                                )
                            )
                        )
                        .ToList();

                    var newWatchedDependencies = newReferences
                        .Concat(newProject.Imports.Select(import => import.ImportedProject.FullPath))
                        .Where(File.Exists)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    var newProjectIncludeDirectories = this.RecalculateProjectIncludeDirectories(newProject);
                    var newPublicIncludeDirectories = this.RecalculatePublicIncludeDirectories(newReferences, globalProperties);

                    Microsoft.Build.Evaluation.Project? oldProject;
                    bool includeEnvironmentChanged;
                    lock (_stateLock) {
                        if (_disposed) {
                            MsBuildProjectAnalyzer.UnloadProject(newProject);
                            return;
                        }

                        includeEnvironmentChanged =
                            !newReferences.SequenceEqual(_currentReferences, StringComparer.OrdinalIgnoreCase) ||
                            !newProjectIncludeDirectories.SequenceEqual(_projectIncludeDirectories, StringComparer.OrdinalIgnoreCase) ||
                            !newPublicIncludeDirectories.SequenceEqual(_publicIncludeDirsFromReferences, StringComparer.OrdinalIgnoreCase);

                        oldProject = _loadedProject;
                        _loadedProject = newProject;
                        _currentReferences = newReferences;
                        _projectIncludeDirectories = newProjectIncludeDirectories;
                        _publicIncludeDirsFromReferences = newPublicIncludeDirectories;
                    }

                    newProject = null;
                    newProjectCollection = null;
                    MsBuildProjectAnalyzer.UnloadProject(oldProject);
                    this.SynchronizeDependencyWatchers(newWatchedDependencies);

                    if (includeEnvironmentChanged) {
                        this.IncludeEnvironmentChanged?.Invoke(_projectFilePath);
                    }
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogError($"[MsBuildProjectAnalyzer] Reload failed for '{_projectFilePath}': {ex}");
                    MsBuildProjectAnalyzer.UnloadProject(newProject);
                    newProjectCollection?.Dispose();
                }
            }
        }

        private List<string> RecalculateProjectIncludeDirectories(Microsoft.Build.Evaluation.Project project) {
            var result = new List<string>();
            string baseDirectory = Path.GetDirectoryName(project.FullPath)!;

            if (project.ItemDefinitions.TryGetValue("ClCompile", out var clCompile)) {
                this.AppendDirectories(result, clCompile.GetMetadataValue("AdditionalIncludeDirectories"), project, baseDirectory);
            }

            // IncludePath содержит стандартные VC/SDK paths и значения из property sheets.
            // ExternalIncludePath используется новыми версиями MSVC для external headers.
            this.AppendDirectories(result, project.GetPropertyValue("IncludePath"), project, baseDirectory);
            this.AppendDirectories(result, project.GetPropertyValue("ExternalIncludePath"), project, baseDirectory);

            return result
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<string> RecalculatePublicIncludeDirectories(
            IReadOnlyList<string> referencePaths,
            IReadOnlyDictionary<string, string> globalProperties
        ) {
            var result = new List<string>();

            foreach (var referencePath in referencePaths) {
                if (!File.Exists(referencePath)) {
                    continue;
                }

                ProjectCollection? projectCollection = null;
                Microsoft.Build.Evaluation.Project? project = null;

                try {
                    projectCollection = new ProjectCollection(
                        globalProperties.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase)
                    );
                    project = projectCollection.LoadProject(referencePath);
                    string baseDirectory = Path.GetDirectoryName(referencePath)!;
                    this.AppendDirectories(result, project.GetPropertyValue("PublicIncludeDirectories"), project, baseDirectory);
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogError($"[MsBuildProjectAnalyzer] Failed to read PublicIncludeDirectories from '{referencePath}': {ex}");
                }
                finally {
                    MsBuildProjectAnalyzer.UnloadProject(project);
                    if (project == null) {
                        projectCollection?.Dispose();
                    }
                }
            }

            return result
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void AppendDirectories(
            List<string> result,
            string rawDirectories,
            Microsoft.Build.Evaluation.Project project,
            string baseDirectory
        ) {
            foreach (var rawPart in rawDirectories.Split(';')) {
                string trimmed = rawPart.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed == "%(AdditionalIncludeDirectories)") {
                    continue;
                }

                try {
                    string expanded = project.ExpandString(trimmed)
                        .Replace("$(MSBuildThisFileDirectory)", baseDirectory)
                        .Replace("$(ProjectDir)", baseDirectory);

                    // Не превращаем нераскрытую MSBuild-переменную в формально корректный,
                    // но заведомо ложный путь относительно каталога проекта.
                    if (expanded.Contains("$(") || expanded.Contains("%(")) {
                        continue;
                    }

                    string fullPath = Path.IsPathRooted(expanded)
                        ? Path.GetFullPath(expanded)
                        : Path.GetFullPath(Path.Combine(baseDirectory, expanded.Replace('/', '\\')));

                    result.Add(fullPath);
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogWarning($"[MsBuildProjectAnalyzer] Skip include directory '{trimmed}' in '{_projectFilePath}': {ex.Message}");
                }
            }
        }

        private void SynchronizeDependencyWatchers(IEnumerable<string> dependencyPaths) {
            var requiredPaths = dependencyPaths
                .Where(File.Exists)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var requiredDirectories = requiredPaths
                .Select(Path.GetDirectoryName)
                .Where(directory => !string.IsNullOrEmpty(directory))
                .Cast<string>()
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            lock (_stateLock) {
                _watchedDependencyPaths = requiredPaths;
            }

            foreach (string staleDirectory in _dependencyDirectoryWatchers.Keys.Where(path => !requiredDirectories.Contains(path)).ToList()) {
                _dependencyDirectoryWatchers[staleDirectory].Dispose();
                _dependencyDirectoryWatchers.Remove(staleDirectory);
            }

            // Один watcher на каталог существенно дешевле, чем отдельный watcher на каждый imported
            // .props/.targets. Фильтрацию до конкретных MSBuild inputs выполняют обработчики ниже.
            foreach (string dependencyDirectory in requiredDirectories.Where(path => !_dependencyDirectoryWatchers.ContainsKey(path))) {
                var watcher = new FileSystemWatcher(dependencyDirectory) {
                    Filter = "*.*",
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName,
                    EnableRaisingEvents = true
                };

                watcher.Changed += this.OnWatchedDependencyChanged;
                watcher.Created += this.OnWatchedDependencyChanged;
                watcher.Renamed += this.OnWatchedDependencyRenamed;
                watcher.Deleted += this.OnWatchedDependencyChanged;
                _dependencyDirectoryWatchers[dependencyDirectory] = watcher;
            }
        }

        private void OnWatchedProjectChanged(object sender, FileSystemEventArgs e) {
            this.ReloadProject();
        }

        private void OnWatchedProjectRenamed(object sender, RenamedEventArgs e) {
            this.ReloadProject();
        }

        private void OnWatchedDependencyChanged(object sender, FileSystemEventArgs e) {
            if (this.IsWatchedDependencyPath(e.FullPath)) {
                this.ReloadProject();
            }
        }

        private void OnWatchedDependencyRenamed(object sender, RenamedEventArgs e) {
            if (this.IsWatchedDependencyPath(e.FullPath) || this.IsWatchedDependencyPath(e.OldFullPath)) {
                this.ReloadProject();
            }
        }

        private bool IsWatchedDependencyPath(string filePath) {
            lock (_stateLock) {
                return _watchedDependencyPaths.Contains(filePath);
            }
        }

        private static void UnloadProject(Microsoft.Build.Evaluation.Project? project) {
            if (project == null) {
                return;
            }

            try {
                ProjectCollection projectCollection = project.ProjectCollection;
                projectCollection.UnloadAllProjects();
                projectCollection.Dispose();
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogWarning($"[MsBuildProjectAnalyzer] Failed to unload MSBuild project: {ex.Message}");
            }
        }

        private static bool EvaluationContextsEqual(
            IReadOnlyDictionary<string, string> left,
            IReadOnlyDictionary<string, string> right
        ) {
            return left.Count == right.Count &&
                   left.All(pair => right.TryGetValue(pair.Key, out string value) && StringComparer.OrdinalIgnoreCase.Equals(pair.Value, value));
        }
    }


    public sealed class MsBuildSolutionWatcher : IDisposable {
        public event Action<string>? IncludeEnvironmentChanged;

        private readonly Dictionary<string, MsBuildProjectAnalyzer> _analyzers = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, EnvDTE.Project> _dteProjects = new(StringComparer.OrdinalIgnoreCase);
        private bool _disposed = false;

        public MsBuildSolutionWatcher(IEnumerable<EnvDTE.Project> projects) {
            ThreadHelper.ThrowIfNotOnUIThread();
            MsBuildEnvironment.EnsureInitialized();
            this.SynchronizeProjects(projects);
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }

            foreach (var analyzer in _analyzers.Values) {
                analyzer.IncludeEnvironmentChanged -= this.OnAnalyzerIncludeEnvironmentChanged;
                analyzer.Dispose();
            }

            _analyzers.Clear();
            _dteProjects.Clear();
            _disposed = true;
        }

        public void SynchronizeProjects(IEnumerable<EnvDTE.Project> projects) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var currentProjectPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var project in projects) {
                string fullPath;
                try {
                    fullPath = Path.GetFullPath(project.FullName);
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogWarning($"[MsBuildSolutionWatcher] Skip project with invalid path: {ex.Message}");
                    continue;
                }

                if (!File.Exists(fullPath)) {
                    continue;
                }

                currentProjectPaths.Add(fullPath);
                this.AddOrUpdateProject(project);
            }

            foreach (string staleProjectPath in _analyzers.Keys.Where(path => !currentProjectPaths.Contains(path)).ToList()) {
                this.RemoveProject(staleProjectPath);
            }
        }

        public void AddOrUpdateProject(EnvDTE.Project project) {
            ThreadHelper.ThrowIfNotOnUIThread();

            string fullPath = Path.GetFullPath(project.FullName);
            if (!File.Exists(fullPath)) {
                return;
            }

            var globalProperties = MsBuildSolutionWatcher.GetActiveGlobalProperties(project);
            _dteProjects[fullPath] = project;

            if (_analyzers.TryGetValue(fullPath, out var analyzer)) {
                analyzer.ReloadProject(globalProperties);
                return;
            }

            analyzer = new MsBuildProjectAnalyzer(fullPath, globalProperties);
            analyzer.IncludeEnvironmentChanged += this.OnAnalyzerIncludeEnvironmentChanged;
            _analyzers[fullPath] = analyzer;
        }

        public void RemoveProject(string projectPath) {
            string fullPath = Path.GetFullPath(projectPath);
            if (_analyzers.TryGetValue(fullPath, out var analyzer)) {
                analyzer.IncludeEnvironmentChanged -= this.OnAnalyzerIncludeEnvironmentChanged;
                analyzer.Dispose();
                _analyzers.Remove(fullPath);
            }

            _dteProjects.Remove(fullPath);
        }

        public IReadOnlyList<string> GetIncludeDirectoriesFor(string projectPath) {
            ThreadHelper.ThrowIfNotOnUIThread();

            string fullPath = Path.GetFullPath(projectPath);
            if (!_analyzers.TryGetValue(fullPath, out var analyzer)) {
                return Array.Empty<string>();
            }

            if (_dteProjects.TryGetValue(fullPath, out var dteProject)) {
                analyzer.EnsureEvaluationContext(MsBuildSolutionWatcher.GetActiveGlobalProperties(dteProject));
            }

            return analyzer.IncludeDirectories;
        }


        public bool RefreshEvaluationContexts() {
            ThreadHelper.ThrowIfNotOnUIThread();

            bool anyContextChanged = false;
            foreach (var pair in _dteProjects.ToList()) {
                if (_analyzers.TryGetValue(pair.Key, out var analyzer)) {
                    anyContextChanged |= analyzer.EnsureEvaluationContext(MsBuildSolutionWatcher.GetActiveGlobalProperties(pair.Value));
                }
            }

            return anyContextChanged;
        }

        public IReadOnlyList<string> GetAllProjectPaths() {
            return _analyzers.Keys.ToList();
        }

        private void OnAnalyzerIncludeEnvironmentChanged(string projectPath) {
            this.IncludeEnvironmentChanged?.Invoke(projectPath);
        }

        private static Dictionary<string, string> GetActiveGlobalProperties(EnvDTE.Project project) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try {
                EnvDTE.Configuration? activeConfiguration = project.ConfigurationManager?.ActiveConfiguration;
                if (activeConfiguration != null) {
                    result["Configuration"] = activeConfiguration.ConfigurationName;
                    result["Platform"] = activeConfiguration.PlatformName;
                }
            }
            catch (Exception ex) {
                // Shared projects и некоторые virtual projects не предоставляют ConfigurationManager.
                Helpers.Diagnostic.Logger.LogDebug($"[MsBuildSolutionWatcher] Project '{project.Name}' has no active configuration: {ex.Message}");
            }

            return result;
        }
    }
}
