using System;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Utilities;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.ComponentModelHost;

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
        private ITextSnapshot? _activeSnapshot;
        private IVsTextView? _activeTextView;

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
                _overlayManager.Overlay.Visibility = Visibility.Visible;
            }
        }

        public void Hide() {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_overlayManager?.Overlay != null) {
                _overlayManager.Overlay.Visibility = Visibility.Collapsed;
            }
        }

        public void Dispose() {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.EnsureDisposed();
        }

        public void ActivateEditorFrame(IVsTextView textView) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // VS может несколько раз сообщить об активации одного frame. Повторный поиск
            // WPF view и пересчёт overlay для него не нужны.
            if (ReferenceEquals(_activeTextView, textView) && _overlayManager?.Overlay?.Visibility == Visibility.Visible) {
                return;
            }

            var viewHost = TextEditorControlHelper.TryGetViewHost(textView);
            if (viewHost == null) {
                // Не показываем данные предыдущего документа, пока новый editor view ещё недоступен.
                this.Hide();
                return;
            }

            _activeTextView = textView;
            _activeSnapshot = viewHost.TextView.TextSnapshot;
            this.Show();
            this.ApplyActiveSnapshot();
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
            if (_overlayManager != null && _overlayManager.IsAttached) {
                return;
            }

            var overlayTargaet = this.TryFindOverlayTarget();
            if (overlayTargaet == null) {
                // Редактор ещё не загружен — повторим позже
                Helpers.Diagnostic.Logger.LogDebug("TextEditor not loaded yet, try later");
                VsixThreadHelper.RunOnUiThread(Dispatcher.CurrentDispatcher, UpdateState, DispatcherPriority.ApplicationIdle);
                return;
            }
            
            var overlay = new Controls.TextEditorOverlayControl();
            _overlayManager = new Helpers.AdornerOverlayManager<Controls.TextEditorOverlayControl>(
                overlayTargaet,
                overlay
                );

            Helpers.Diagnostic.Logger.LogDebug("AdornerOverlayManager created");
        }

        private void ApplyActiveSnapshot() {
            if (_activeSnapshot != null && _overlayManager?.Overlay != null) {
                _overlayManager.Overlay.OnActiveTextViewChanged(_activeSnapshot);
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
            _activeTextView = null;
            _activeSnapshot = null;

            Helpers.Diagnostic.Logger.LogDebug("AdornerOverlayManager disposed");
        }


        /// <summary>
        /// Ищет WpfTextViewHost и возвращает его родителя с именем PART_ContentPanel.
        /// </summary>
        private FrameworkElement TryFindOverlayTarget() {
            // NOTE: Элемент с именем "PART_ContentPanel" не уникален, поэтому сначала
            //       ищем уникальный элемент текстового редаквтора - WpfTextViewHost, а затем его родителя.
            var viewHost = Helpers.VisualTree.FindChildByType(Application.Current.MainWindow, "WpfTextViewHost");
            if (viewHost == null) {
                return null;
            }

            var panel = Helpers.VisualTree.FindParentByName(viewHost, "PART_ContentPanel");
            return panel;
        }
    }
}
