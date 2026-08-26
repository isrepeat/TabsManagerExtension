using System;
using System.Linq;
using System.Windows.Threading;
using Microsoft.VisualStudio.Shell;

using TMEx = TabsManagerExtension;

namespace TabsManagerExtension.Controls.Tabs {
    // Синхронизирует модели вкладок с жизненным циклом solution, DTE-документов,
    // tool window и файлов на диске. Здесь же формируется история закрытых вкладок.
    internal sealed class TabsWorkspaceSynchronizer {
        private readonly EnvDTE80.DTE2 _dte;
        private readonly Dispatcher _dispatcher;
        private readonly TabCollectionManager _tabCollectionManager;
        private readonly ClosedTabsHistory _closedTabsHistory;
        private readonly VsShell.TextEditor.Overlay.TextEditorOverlayController _overlayController;
        private readonly ToolWindowSessionManager _toolWindowSessionManager;
        private readonly TabActivationSynchronizer _activationSynchronizer;
        private readonly Action _onStopFileMonitor;

        private TabsStateReconciler? _stateReconciler;
        private string? _loadedSolutionName;

        public TabsWorkspaceSynchronizer(
            EnvDTE80.DTE2 dte,
            Dispatcher dispatcher,
            TabCollectionManager tabCollectionManager,
            ClosedTabsHistory closedTabsHistory,
            VsShell.TextEditor.Overlay.TextEditorOverlayController overlayController,
            ToolWindowSessionManager toolWindowSessionManager,
            TabActivationSynchronizer activationSynchronizer,
            Action onStopFileMonitor
            ) {
            _dte = dte;
            _dispatcher = dispatcher;
            _tabCollectionManager = tabCollectionManager;
            _closedTabsHistory = closedTabsHistory;
            _overlayController = overlayController;
            _toolWindowSessionManager = toolWindowSessionManager;
            _activationSynchronizer = activationSynchronizer;
            _onStopFileMonitor = onStopFileMonitor;
        }

        public void SetStateReconciler(TabsStateReconciler stateReconciler) {
            _stateReconciler = stateReconciler;
        }

        public bool EnsureSolutionLoaded(string solutionName, Action onEnsureFileMonitor) {
            // InitialAnalysisCompleted может повторно уведомить уже загруженный solution.
            // Не очищаем коллекции во второй раз и не запускаем дублирующий file monitor.
            if (string.Equals(_loadedSolutionName, solutionName, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            _loadedSolutionName = solutionName;
            onEnsureFileMonitor();
            this.LoadOpenTabs();
            return true;
        }

        public void HandleDocumentOpened(EnvDTE.Document document) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // События DTE иногда приходят после того, как модель уже была создана обработчиком
            // активации. Повторно используем её и только уточняем фактическую группу документа.
            var tabItemDocument = _tabCollectionManager.Find(document) ?? new TMEx.State.Document.TabItemDocument(document);
            if (tabItemDocument.ShellDocument.IsDocumentInPreviewTab()) {
                _tabCollectionManager.AddDocumentToPreview(tabItemDocument);
            }
            else {
                _tabCollectionManager.AddAutomatically(tabItemDocument);
            }
        }

        public void HandleDocumentSaved(EnvDTE.Document document) {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Сохранение может изменить preview-состояние или данные, получаемые от VS Shell.
            _stateReconciler?.Reconcile();
        }

        public void HandleDocumentClosing(EnvDTE.Document document) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (document == null) {
                return;
            }

            if (!string.IsNullOrEmpty(document.FullName)) {
                // Overlay должен отпустить закрываемый editor frame до удаления модели вкладки.
                _overlayController.OnDocumentClosing(document.FullName);
            }

            var tabItemDocument = _tabCollectionManager.Find(document);
            if (tabItemDocument == null) {
                Helpers.Diagnostic.Logger.LogWarning($"\"{document.Name}\" not found in collections");
                return;
            }

            // Командное закрытие заранее ставит closing mark и само объединяет вкладки в одну
            // undo-операцию. Немаркированное DTE-событие считаем внешним одиночным закрытием.
            if (!VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.IsSolutionClosing &&
                !_closedTabsHistory.ConsumeClosingMark(tabItemDocument)) {
                _closedTabsHistory.Push(new[] { this.CreateClosedTabEntry(tabItemDocument) });
            }

            // DocumentClosing вызывается синхронно из DTE Close(), поэтому удаление здесь
            // сразу становится видимым коду, который проверяет результат команды закрытия.
            _tabCollectionManager.Remove(tabItemDocument);
        }

        public void HandleWindowClosing(EnvDTE.Window closingWindow) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var tabItemWindow = _tabCollectionManager.Find(closingWindow);
            if (tabItemWindow != null &&
                !VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.IsSolutionClosing &&
                !_closedTabsHistory.ConsumeClosingMark(tabItemWindow)) {

                _closedTabsHistory.Push(new[] { this.CreateClosedTabEntry(tabItemWindow) });
            }

            // WindowClosing приходит до окончательного обновления DTE.Windows. Откладываем
            // сохранение, чтобы снимок tool window отражал уже завершённое закрытие.
            VsixThreadHelper.RunOnUiThread(
                _dispatcher,
                () => _toolWindowSessionManager.Save(VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.IsSolutionClosing),
                DispatcherPriority.Background
            );
        }

        public void HandleSolutionClosing() {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Сохраняем состояние до массовых WindowClosing/DocumentClosing и запрещаем этим
            // событиям заполнять undo-историю вкладками закрываемого solution.
            // Сервис уже выставил IsSolutionClosing перед публикацией события. Здесь намеренно
            // сохраняем последний рабочий снимок, прежде чем начнут закрываться документы.
            _toolWindowSessionManager.Save(isSolutionClosing: false);
            _toolWindowSessionManager.CancelActiveWindowRestore();
        }

        public void HandleSolutionClosed(string solutionName) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Полная очистка выполняется только после подтверждённого закрытия. До этого
            // состояние нужно сохранить на случай отмены диалога сохранения документов.
            Helpers.Diagnostic.Logger.LogDebug($"Solution closed; clearing Tabs Manager state for '{solutionName}'.");
            _closedTabsHistory.Clear();
            _tabCollectionManager.Clear();
            _onStopFileMonitor();
            _loadedSolutionName = null;
        }

        public void HandleFileChanged(string fullPath) {
            // FileSystemWatcher может сообщить изменение без DTE-события. Обновление пути
            // без нового имени заставляет модель перечитать вычисляемые данные документа.
            _tabCollectionManager.UpdateDocumentPath(fullPath);
        }

        public void HandleFileRenamed(string oldFullPath, string newFullPath) {
            // Сохраняем существующий TabItem и его selection/group вместо удаления и создания.
            _tabCollectionManager.UpdateDocumentPath(oldFullPath, newFullPath);
        }

        public void HandleFileDeleted(string fullPath) {
            var tabItemDocument = _tabCollectionManager.Find(fullPath);
            if (tabItemDocument != null) {
                _tabCollectionManager.Remove(tabItemDocument);
            }
        }

        public void UpdateWindowTabsInfo() {
            ThreadHelper.ThrowIfNotOnUIThread();

            // WindowId устойчивее Caption и позволяет сопоставить сохранённый TabItemWindow
            // с текущим экземпляром DTE.Window после восстановления сессии.
            var windowsById = _dte.Windows
                .Cast<EnvDTE.Window>()
                .Select(window => new {
                    Window = window,
                    Id = VsShell.Document.ShellWindow.GetWindowId(window)
                })
                .Where(entry => !string.IsNullOrEmpty(entry.Id))
                .GroupBy(entry => entry.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Window, StringComparer.OrdinalIgnoreCase);

            _tabCollectionManager.ForEach<TMEx.State.Document.TabItemWindow>(tabItemWindow => {
                try {
                    if (windowsById.TryGetValue(tabItemWindow.WindowId, out var matchingWindow) &&
                        tabItemWindow.Caption != matchingWindow.Caption) {

                        tabItemWindow.Caption = matchingWindow.Caption;
                        tabItemWindow.FullName = matchingWindow.Caption;
                    }
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogError($"Failed to update caption for TabItemWindow: {ex.Message}");
                }
            });
        }

        private void LoadOpenTabs() {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Solution загружается как единый снимок: сначала сохранённые tool window, затем
            // открытые документы и ещё не представленные окна текущего экземпляра VS.
            _tabCollectionManager.Clear();
            _toolWindowSessionManager.Restore();
            foreach (EnvDTE.Document document in _dte.Documents) {
                _tabCollectionManager.AddAutomatically(new TMEx.State.Document.TabItemDocument(document));
            }
            foreach (EnvDTE.Window window in _dte.Windows) {
                if (window.Document != null) {
                    continue;
                }

                var shellWindow = new VsShell.Document.ShellWindow(window);
                if (shellWindow.IsTabWindow()) {
                    _tabCollectionManager.AddAutomatically(new TMEx.State.Document.TabItemWindow(shellWindow));
                }
            }

            _activationSynchronizer.SyncWithActiveWindow();
            // ActiveWindow и активный editor view стабилизируются позже перечисления DTE.
            // Поэтому начальное состояние overlay вычисляем в фоновой UI-очереди.
            VsixThreadHelper.RunOnUiThread(_dispatcher, () => {
                var activeWindow = _dte.ActiveWindow;
                bool isDocumentFrameActive = activeWindow?.Document != null &&
                    VsShell.Document.ShellWindow.IsTabWindow(activeWindow);
                if (isDocumentFrameActive && VsShell.TextEditor.TextEditorControlHelper.IsEditorActive()) {
                    Helpers.GlobalFlags.SetFlag("TextEditorFrameFocused", true);
                    _overlayController.Show();
                }
                else {
                    Helpers.GlobalFlags.SetFlag("TextEditorFrameFocused", false);
                    _overlayController.Hide();
                }
            }, DispatcherPriority.Background);
        }

        public ClosedTabEntry CreateClosedTabEntry(TMEx.State.Document.TabItemBase tabItem) {
            // Группа фиксируется до синхронного удаления TabItem из коллекции и используется
            // при восстановлении для возврата вкладки на прежнее место.
            var current = _tabCollectionManager.FindWithGroup(tabItem);
            return _closedTabsHistory.CreateEntry(tabItem, current?.Group);
        }
    }
}
