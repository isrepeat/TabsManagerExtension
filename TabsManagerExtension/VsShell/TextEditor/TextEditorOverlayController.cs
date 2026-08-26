using System;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TextManager.Interop;

namespace TabsManagerExtension.VsShell.TextEditor.Overlay {

    /// <summary>
    /// Отвечает за управление жизненным циклом визуального оверлея (`TextEditorOverlayControl`),
    /// который добавляется поверх текстового редактора через `AdornerOverlayManager`.
    /// 
    /// Контроллер следит за количеством открытых документов через DTE и:
    /// - создаёт оверлей при появлении хотя бы одного редактора,
    /// - удаляет оверлей, когда все редакторы закрыты.
    /// </summary>
    public class TextEditorOverlayController {
        private readonly EnvDTE80.DTE2 _dte2;

        private Helpers.AdornerOverlayManager<Controls.TextEditorOverlayControl>? _overlayManager;
        private Controls.TextEditorOverlayControl? _overlay;
        private FrameworkElement? _overlayTarget;
        private ITextSnapshot? _activeSnapshot;
        private IVsTextView? _activeTextView;
        private IWpfTextView? _activeWpfTextView;
        private string? _activeDocumentFullName;

        /// <summary>
        /// Инициализирует контроллер, привязанный к текущему экземпляру Visual Studio (DTE).
        /// </summary>
        public TextEditorOverlayController(EnvDTE80.DTE2 dte2) {
            _dte2 = dte2;
        }


        public void Show() {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.EnsureCreated();

            if (_overlayManager?.Overlay != null) {
                _overlayManager.Overlay.OnEditorFrameActivityChanged(
                    isActive: true,
                    keepFindCommandsWhenInactive: false
                );

                _overlayManager.Overlay.Visibility = Visibility.Visible;
            }
        }

        public void Hide() {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_overlayManager?.Overlay != null) {
                _overlayManager.Overlay.OnEditorFrameActivityChanged(
                    isActive: false,
                    keepFindCommandsWhenInactive: false
                );

                _overlayManager.Overlay.Visibility = Visibility.Collapsed;
            }
        }

        public void DeactivateEditorFrame() {
            ThreadHelper.ThrowIfNotOnUIThread();
            // При переходе фокуса в tool window оставляем доступными команды открытой Quick Find.
            _overlayManager?.Overlay?.OnEditorFrameActivityChanged(
                isActive: false,
                keepFindCommandsWhenInactive: true
            );
        }

        public void Dispose() {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.EnsureDisposed();
        }

        public void ActivateEditorFrame(IVsTextView textView) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var viewHost = TextEditorControlHelper.TryGetViewHost(textView);
            if (viewHost == null) {
                // Не показываем данные предыдущего документа, пока новый editor view ещё недоступен.
                this.Hide();
                return;
            }

            var overlayTarget = this.TryFindOverlayTarget(viewHost.TextView.VisualElement);
            if (overlayTarget == null) {
                // Не оставляем overlay предыдущего документа поверх нового editor frame.
                this.Hide();
                return;
            }

            _activeTextView = textView;
            _activeWpfTextView = viewHost.TextView;
            _activeSnapshot = viewHost.TextView.TextSnapshot;
            _activeDocumentFullName = _dte2.ActiveDocument?.FullName;
            this.EnsureAttached(overlayTarget);
            if (_overlay != null) {
                _overlay.OnEditorFrameActivityChanged(
                    isActive: true,
                    keepFindCommandsWhenInactive: false
                );

                _overlay.Visibility = Visibility.Visible;
            }

            this.ApplyActiveSnapshot();
        }

        public void OnDocumentClosing(string documentFullName) {
            ThreadHelper.ThrowIfNotOnUIThread();

            bool isTrackedDocument = string.Equals(
                _activeDocumentFullName,
                documentFullName,
                StringComparison.OrdinalIgnoreCase
            );

            bool isDteActiveDocument = string.Equals(
                _dte2.ActiveDocument?.FullName,
                documentFullName,
                StringComparison.OrdinalIgnoreCase
            );

            if (!isTrackedDocument && !isDteActiveDocument) {
                return;
            }

            _overlay?.ResetClosedDocumentState();
            _activeSnapshot = null;
            _activeTextView = null;
            _activeWpfTextView = null;
            _activeDocumentFullName = null;
        }



        /// <summary>
        /// Обновляет состояние оверлея:
        /// - создаёт, если есть хотя бы один открытый документ;
        /// - уничтожает, если все документы закрыты.
        /// </summary>
        public void UpdateState() {
            ThreadHelper.ThrowIfNotOnUIThread();

            bool hasOpenDocuments = _dte2.Documents.Cast<EnvDTE.Document>().Any();
            if (hasOpenDocuments) {
                this.EnsureCreated();
                this.ApplyActiveSnapshot();
            }
            else {
                this.EnsureDisposed();
            }
        }

        /// <summary>
        /// Создаёт визуальный оверлей, если он ещё не создан или был откреплён.
        /// </summary>
        private void EnsureCreated() {
            var viewHost = TextEditorControlHelper.TryGetActiveViewHost();
            var overlayTarget = viewHost == null
                ? null
                : this.TryFindOverlayTarget(viewHost.TextView.VisualElement);
            if (overlayTarget == null) {
                // Редактор ещё не загружен — повторим позже
                Helpers.Diagnostic.Logger.LogDebug("TextEditor not loaded yet, try later");
                VsixThreadHelper.RunOnUiThread(Dispatcher.CurrentDispatcher, UpdateState, DispatcherPriority.ApplicationIdle);
                return;
            }

            // Show может прийти из DTE WindowActivated без события IVsWindowFrame.
            // Поэтому snapshot активного view синхронизируем и в этом пути.
            _activeSnapshot = viewHost.TextView.TextSnapshot;
            _activeWpfTextView = viewHost.TextView;
            _activeDocumentFullName = _dte2.ActiveDocument?.FullName;
            this.EnsureAttached(overlayTarget);
            this.ApplyActiveSnapshot();
        }

        private void EnsureAttached(FrameworkElement overlayTarget) {
            if (_overlayManager?.IsAttached == true && ReferenceEquals(_overlayTarget, overlayTarget)) {
                return;
            }

            // У разных document frame собственный PART_ContentPanel. Переносим один и тот же
            // контрол, чтобы при переключении редакторов сохранить состояние и кэш якорей.
            _overlayManager?.Remove();
            _overlay ??= new Controls.TextEditorOverlayControl();
            _overlayManager = new Helpers.AdornerOverlayManager<Controls.TextEditorOverlayControl>(
                overlayTarget,
                _overlay
            );

            _overlayTarget = overlayTarget;

            Helpers.Diagnostic.Logger.LogDebug("AdornerOverlayManager created");
        }

        private void ApplyActiveSnapshot() {
            if (_activeWpfTextView != null && _overlayManager?.Overlay != null) {
                _overlayManager.Overlay.OnActiveTextViewChanged(_activeWpfTextView);
            }
        }

        /// <summary>
        /// Удаляет визуальный оверлей, если он существует.
        /// </summary>
        private void EnsureDisposed() {
            if (_overlayManager == null) {
                return;
            }

            _overlayManager.Remove();
            _overlayManager = null;
            _overlay = null;
            _overlayTarget = null;
            _activeTextView = null;
            _activeWpfTextView = null;
            _activeSnapshot = null;
            _activeDocumentFullName = null;

            Helpers.Diagnostic.Logger.LogDebug("AdornerOverlayManager disposed");
        }


        /// <summary>
        /// Ищет WpfTextViewHost и возвращает его родителя с именем PART_ContentPanel.
        /// </summary>
        private FrameworkElement? TryFindOverlayTarget(FrameworkElement textViewVisualElement) {
            // PART_ContentPanel не уникален, поэтому поднимаемся строго от visual element
            // активного IWpfTextView, а не выбираем первый редактор в главном окне VS.
            return Helpers.VisualTree.FindParentByName(textViewVisualElement, "PART_ContentPanel");
        }
    }
}
