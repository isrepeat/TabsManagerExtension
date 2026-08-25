using System;
using System.IO;
using System.Windows.Threading;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Наблюдает за файлами solution и маршалит события FileSystemWatcher в WPF Dispatcher.</summary>
    internal sealed class SolutionFileChangeMonitor : IDisposable {
        private readonly Dispatcher _dispatcher;

        private FileSystemWatcher? _fileWatcher;

        public event Action<string>? FileChanged;
        public event Action<string, string>? FileRenamed;
        public event Action<string>? FileDeleted;

        public bool IsRunning => _fileWatcher != null;

        public SolutionFileChangeMonitor(Dispatcher dispatcher) {
            _dispatcher = dispatcher;
        }

        public void Start(string? solutionFullName) {
            this.Stop();

            string? solutionDirectory = string.IsNullOrEmpty(solutionFullName)
                ? null
                : Path.GetDirectoryName(solutionFullName);

            if (string.IsNullOrEmpty(solutionDirectory)) {
                return;
            }

            var fileWatcher = new FileSystemWatcher {
                Path = solutionDirectory,
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size,
                IncludeSubdirectories = true,
                Filter = "*.*"
            };

            fileWatcher.Changed += this.OnFileChanged;
            fileWatcher.Renamed += this.OnFileRenamed;
            fileWatcher.Deleted += this.OnFileDeleted;
            _fileWatcher = fileWatcher;
            fileWatcher.EnableRaisingEvents = true;
        }

        public void Stop() {
            var fileWatcher = _fileWatcher;
            _fileWatcher = null;
            if (fileWatcher == null) {
                return;
            }

            try {
                fileWatcher.EnableRaisingEvents = false;
                fileWatcher.Changed -= this.OnFileChanged;
                fileWatcher.Renamed -= this.OnFileRenamed;
                fileWatcher.Deleted -= this.OnFileDeleted;
                fileWatcher.Dispose();
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Error while disposing FileSystemWatcher: {ex}");
            }
        }

        public void Dispose() {
            this.Stop();
        }

        private void OnFileChanged(object sender, FileSystemEventArgs e) {
            if (!IsTemporaryFile(e.FullPath)) {
                this.Schedule(sender, () => this.FileChanged?.Invoke(e.FullPath));
            }
        }

        private void OnFileRenamed(object sender, RenamedEventArgs e) {
            if (!IsTemporaryFile(e.FullPath) && !IsTemporaryFile(e.OldFullPath)) {
                this.Schedule(sender, () => this.FileRenamed?.Invoke(e.OldFullPath, e.FullPath));
            }
        }

        private void OnFileDeleted(object sender, FileSystemEventArgs e) {
            if (!IsTemporaryFile(e.FullPath)) {
                this.Schedule(sender, () => this.FileDeleted?.Invoke(e.FullPath));
            }
        }

        private void Schedule(object sender, Action update) {
            // FileSystemWatcher вызывает обработчики на фоновом потоке. Переход на UI-поток
            // выполняется асинхронно, чтобы не создать взаимную блокировку при закрытии VS.
            if (_dispatcher.HasShutdownStarted || _dispatcher.HasShutdownFinished) {
                return;
            }

            try {
                _dispatcher.BeginInvoke(
                    new Action(() => {
                        if (ReferenceEquals(sender, _fileWatcher)) {
                            update();
                        }
                    }),
                    DispatcherPriority.Background
                );
            }
            catch (InvalidOperationException) {
                // Dispatcher завершился между проверкой и постановкой действия в очередь.
            }
        }

        private static bool IsTemporaryFile(string fullPath) {
            string extension = Path.GetExtension(fullPath);
            return extension.Equals(".TMP", StringComparison.OrdinalIgnoreCase) ||
                fullPath.Contains("~") && fullPath.Contains(".TMP");
        }
    }
}
