using System;
using System.Linq;
using System.Windows.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Сохраняет и восстанавливает tool window, включая отложенную активацию frame.</summary>
    internal sealed class ToolWindowSessionManager : IDisposable {
        private readonly EnvDTE80.DTE2 _dte;
        private readonly Dispatcher _dispatcher;

        private DispatcherTimer? _activeWindowRestoreTimer;
        private IVsWindowFrame? _pendingActiveWindowFrame;
        private string? _pendingActiveWindowId;
        private bool _isEnforcingPendingWindow;

        public bool IsRestoring { get; private set; }

        public ToolWindowSessionManager(EnvDTE80.DTE2 dte, Dispatcher dispatcher) {
            _dte = dte;
            _dispatcher = dispatcher;
        }

        public void Restore() {
            ThreadHelper.ThrowIfNotOnUIThread();

            var uiShell = Package.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
            if (uiShell == null) {
                return;
            }

            this.IsRestoring = true;
            try {
                string? activeWindowId = Settings.TabsManagerSettingsService.ActiveToolWindowId;
                IVsWindowFrame? activeWindowFrame = null;

                // Сначала восстанавливаем весь набор окон без активации. Активный frame
                // возвращается отдельно после загрузки solution, чтобы не спорить с VS за фокус.
                foreach (var windowId in Settings.TabsManagerSettingsService.OpenToolWindowIds) {
                    if (!Guid.TryParse(windowId, out var persistenceGuid)) {
                        continue;
                    }

                    bool isAlreadyVisible = _dte.Windows
                        .Cast<EnvDTE.Window>()
                        .Any(window =>
                            window.Visible &&
                            window.Document == null &&
                            VsShell.Document.ShellWindow.IsTabWindow(window) &&
                            string.Equals(VsShell.Document.ShellWindow.GetWindowId(window), windowId, StringComparison.OrdinalIgnoreCase)
                        );

                    try {
                        // Persistence GUID позволяет VS найти или создать зарегистрированный
                        // tool window независимо от изменяемого Caption.
                        int result = uiShell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fForceCreate, ref persistenceGuid, out var frame);
                        if (!ErrorHandler.Succeeded(result)) {
                            continue;
                        }

                        if (string.Equals(windowId, activeWindowId, StringComparison.OrdinalIgnoreCase)) {
                            activeWindowFrame = frame;
                        }

                        if (!isAlreadyVisible) {
                            // ShowNoActivate восстанавливает layout, не отбирая фокус у документа,
                            // который Visual Studio открывает в этот же момент.
                            frame?.ShowNoActivate();
                        }
                    }
                    catch (Exception ex) {
                        Helpers.Diagnostic.Logger.LogWarning($"Failed to restore tool window '{windowId}': {ex.Message}");
                    }
                }

                if (activeWindowFrame != null && !string.IsNullOrEmpty(activeWindowId)) {
                    this.ScheduleActiveWindowRestore(activeWindowFrame, activeWindowId!);
                }
            }
            finally {
                this.IsRestoring = false;
            }
        }

        public void PrepareActiveWindowRestore() {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Этот ранний путь вызывается до обычного Restore: frame создаётся заранее, чтобы
            // VS не выбрала другой документ как окончательно активный при загрузке solution.
            string? activeWindowId = Settings.TabsManagerSettingsService.ActiveToolWindowId;
            if (string.IsNullOrEmpty(activeWindowId) || !Guid.TryParse(activeWindowId, out var persistenceGuid)) {
                return;
            }

            var uiShell = Package.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
            if (uiShell == null) {
                return;
            }

            int result = uiShell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fForceCreate, ref persistenceGuid, out var frame);
            if (ErrorHandler.Succeeded(result) && frame != null) {
                Helpers.Diagnostic.Logger.LogDebug($"Preparing active tool window frame before solution restore ({activeWindowId}).");
                this.ScheduleActiveWindowRestore(frame, activeWindowId!);
            }
        }

        public bool KeepPendingWindowActive() {
            ThreadHelper.ThrowIfNotOnUIThread();

            var frame = _pendingActiveWindowFrame;
            if (frame == null || _isEnforcingPendingWindow) {
                return false;
            }

            try {
                // Пока отложенное восстановление не завершено, промежуточную активацию документа
                // компенсируем повторным Show сохранённого frame. Guard предотвращает рекурсию,
                // поскольку Show сам порождает очередное событие активации.
                _isEnforcingPendingWindow = true;
                Helpers.Diagnostic.Logger.LogDebug($"Ignoring document activation while restoring tool window ({_pendingActiveWindowId}).");
                frame.Show();
            }
            finally {
                _isEnforcingPendingWindow = false;
            }

            return true;
        }

        public void Save(bool isSolutionClosing) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Во время Restore DTE содержит неполный набор окон, а при закрытии solution окна
            // массово исчезают. Сохранение такого снимка затёрло бы последнюю рабочую сессию.
            if (this.IsRestoring || isSolutionClosing) {
                return;
            }

            var windowIds = _dte.Windows
                .Cast<EnvDTE.Window>()
                // Документные окна восстанавливает сама solution; здесь сохраняются только
                // видимые tool windows, которые Visual Studio считает tab window.
                .Where(window => window.Visible && window.Document == null && VsShell.Document.ShellWindow.IsTabWindow(window))
                .Select(VsShell.Document.ShellWindow.GetWindowId)
                .ToList();

            var activeWindow = _dte.ActiveWindow;
            string? activeWindowId = activeWindow != null &&
                activeWindow.Document == null &&
                VsShell.Document.ShellWindow.IsTabWindow(activeWindow)
                    ? VsShell.Document.ShellWindow.GetWindowId(activeWindow)
                    : null;

            Settings.TabsManagerSettingsService.SetOpenToolWindowState(windowIds, activeWindowId);
        }

        public void CancelActiveWindowRestore() {
            // Timer владеет подпиской на manager; отписываемся явно, чтобы закрытый control
            // не удерживался dispatcher-очередью и не получил поздний callback.
            if (_activeWindowRestoreTimer != null) {
                _activeWindowRestoreTimer.Stop();
                _activeWindowRestoreTimer.Tick -= this.OnActiveWindowRestoreTimerTick;
                _activeWindowRestoreTimer = null;
            }

            _pendingActiveWindowFrame = null;
            _pendingActiveWindowId = null;
        }

        public void Dispose() {
            this.CancelActiveWindowRestore();
        }

        private void ScheduleActiveWindowRestore(IVsWindowFrame frame, string windowId) {
            ThreadHelper.ThrowIfNotOnUIThread();

            _pendingActiveWindowFrame = frame;
            _pendingActiveWindowId = windowId;
            if (_activeWindowRestoreTimer != null) {
                // Один timer обслуживает последнее запрошенное окно: поля выше намеренно
                // обновлены до проверки, чтобы новый restore заменил устаревшую цель.
                return;
            }

            // Интервал здесь не является фиксированной задержкой готовности. Timer повторяет
            // проверку состояния VS и активирует frame только после выполнения условий ниже.
            _activeWindowRestoreTimer = new DispatcherTimer(DispatcherPriority.ApplicationIdle, _dispatcher) {
                Interval = TimeSpan.FromMilliseconds(250)
            };
            _activeWindowRestoreTimer.Tick += this.OnActiveWindowRestoreTimerTick;
            _activeWindowRestoreTimer.Start();
        }

        private void OnActiveWindowRestoreTimerTick(object sender, EventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!_dte.Solution.IsOpen) {
                // Оставляем timer активным: solution может находиться между началом открытия
                // и установкой DTE.Solution.IsOpen.
                return;
            }

            var solution = PackageServices.TryGetVsSolution();
            if (solution == null || !GetSolutionBooleanProperty(solution, (int)__VSPROPID4.VSPROPID_IsSolutionFullyLoaded)) {
                return;
            }
            if (GetSolutionBooleanProperty(solution, (int)__VSPROPID2.VSPROPID_IsSolutionOpeningDocs)) {
                // FullyLoaded может установиться раньше завершения восстановления документов;
                // ранний Show в этот момент снова будет перекрыт активацией editor frame.
                return;
            }

            var frame = _pendingActiveWindowFrame;
            string? windowId = _pendingActiveWindowId;
            // Сначала останавливаем timer и очищаем pending state: frame.Show() синхронно
            // порождает activation events, которые не должны запустить повторное удержание.
            this.CancelActiveWindowRestore();
            if (frame == null) {
                return;
            }

            Helpers.Diagnostic.Logger.LogDebug($"Restoring active tool window frame after solution load ({windowId}).");
            frame.Show();
        }

        private static bool GetSolutionBooleanProperty(IVsSolution solution, int propertyId) {
            if (ErrorHandler.Failed(solution.GetProperty(propertyId, out var value))) {
                return false;
            }

            // Разные реализации VS возвращают shell-свойства как bool либо COM-style int.
            return value is bool booleanValue
                ? booleanValue
                : value is int integerValue && integerValue != 0;
        }
    }
}
