using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using TabsManagerExtension.State.TextEditor;


namespace TabsManagerExtension.Controls {
    public partial class TextEditorOverlayControl : Helpers.BaseUserControl {
        private Helpers.Properties.VisibilityProperty _isAnchorToggleButtonVisible = new();
        public Helpers.Properties.VisibilityProperty IsAnchorToggleButtonVisible {
            get => _isAnchorToggleButtonVisible;
            set {
                if (_isAnchorToggleButtonVisible != value) {
                    _isAnchorToggleButtonVisible = value;
                    this.OnPropertyChanged();
                }
            }
        }

        private Helpers.Properties.VisibilityProperty _isAnchorListVisible = new();
        public Helpers.Properties.VisibilityProperty IsAnchorListVisible {
            get => _isAnchorListVisible;
            set {
                if (_isAnchorListVisible != value) {
                    _isAnchorListVisible = value;
                    this.OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<State.TextEditor.AnchorPoint> Anchors { get; } = new();

        private State.TextEditor.AnchorPoint? _selectedAnchor;
        private ITextSnapshot? _activeSnapshot;
        private IWpfTextView? _activeTextView;
        private ITextBuffer? _loadedTextBuffer;
        private int _loadedSnapshotVersion = -1;
        private bool _isAnchorListExpanded;
        private bool _isSynchronizingCaretSelection;
        private readonly Dictionary<ITextBuffer, AnchorSnapshotCache> _anchorCache = new();

        private sealed class AnchorSnapshotCache {
            public int Version { get; }
            public bool HasAnchors { get; }
            public IReadOnlyList<State.TextEditor.AnchorPoint>? ParsedAnchors { get; }

            public AnchorSnapshotCache(
                int version,
                bool hasAnchors,
                IReadOnlyList<State.TextEditor.AnchorPoint>? parsedAnchors
                ) {
                this.Version = version;
                this.HasAnchors = hasAnchors;
                this.ParsedAnchors = parsedAnchors;
            }
        }
        public State.TextEditor.AnchorPoint? SelectedAnchor {
            get => _selectedAnchor;
            set {
                if (_selectedAnchor != value) {
                    _selectedAnchor = value;
                    this.OnPropertyChanged();
                    if (value != null && !_isSynchronizingCaretSelection) {
                        this.NavigateToLine(value.LineNumber);
                    }
                }
            }
        }

        public ICommand OnToggleAnchorListCommand { get; }

        public TextEditorOverlayControl() {
            this.InitializeComponent();
            this.Loaded += this.OnLoaded;
            this.Unloaded += this.OnUnloaded;
            this.DataContext = this;

            this.OnToggleAnchorListCommand = new Helpers.RelayCommand<object>(this.OnToggleAnchorList);
        }

        private void OnLoaded(object sender, RoutedEventArgs e) {
            //VsShell.TextEditor.Services.TextEditorCommandFilterService.Instance.AddTrackedInputElement(this);

            // IsHitTestVisible могут быть унаследованы от родителя (например, AdornerLayer),
            // поэтому значения из XAML не применяются гарантированно — устанавливаем явно в OnLoaded.
            this.IsHitTestVisible = true;

        }

        private void OnUnloaded(object sender, RoutedEventArgs e) {
            //VsShell.TextEditor.Services.TextEditorCommandFilterService.Instance.RemoveTrackedInputElement(this);
            // Контрол временно выгружается при переносе adorner между document frame.
            // Не очищаем состояние и кэш: новый snapshot будет применён сразу после повторного подключения.
        }


        private void OnToggleAnchorList(object parameter) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnToggleAnchorList()");

            if (_isAnchorListExpanded) {
                _isAnchorListExpanded = false;
                this.UnsubscribeFromCaret();
                this.IsAnchorListVisible.Value = false;
                this.SelectedAnchor = null;
                return;
            }

            // Пока список закрыт, содержимое не пересчитывается при смене документа.
            // Загружаем anchors лениво только перед фактическим раскрытием.
            _isAnchorListExpanded = true;
            this.SubscribeToCaret();
            if (_activeSnapshot != null) {
                this.LoadAnchors(_activeSnapshot);
            }
            else {
                this.LoadAnchorsFromActiveDocument();
            }
        }


        public void OnActiveTextViewChanged(IWpfTextView textView) {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.SetActiveTextView(textView);
            var snapshot = textView.TextSnapshot;
            _activeSnapshot = snapshot;

            if (!_isAnchorListExpanded) {
                // Для закрытого списка достаточно быстро определить наличие хотя бы одного маркера.
                // Полный разбор документа откладываем до раскрытия списка.
                this.PrepareCollapsedState(snapshot);
                return;
            }

            this.LoadAnchors(snapshot);
        }

        public void ResetClosedDocumentState() {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Открытость списка относится к текущему экземпляру вкладки и не должна
            // переживать закрытие файла с последующим повторным открытием.
            if (_activeSnapshot != null) {
                _anchorCache.Remove(_activeSnapshot.TextBuffer);
            }

            _isAnchorListExpanded = false;
            this.UnsubscribeFromCaret();
            _activeTextView = null;
            this.ResetLoadedAnchors(showToggleButton: false);
        }

        public void LoadAnchorsFromActiveDocument() {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Берём текст напрямую из буфера редактора, чтобы не читать документ построчно через медленный DTE/COM API.
            var viewHost = VsShell.TextEditor.TextEditorControlHelper.TryGetActiveViewHost();
            var snapshot = viewHost?.TextView.TextSnapshot;
            if (snapshot == null) {
                this.ResetLoadedAnchors(showToggleButton: false);
                return;
            }

            _activeSnapshot = snapshot;
            if (viewHost != null) {
                this.SetActiveTextView(viewHost.TextView);
            }

            this.LoadAnchors(snapshot);
        }

        private void LoadAnchors(ITextSnapshot snapshot) {
            ThreadHelper.ThrowIfNotOnUIThread();

            _loadedTextBuffer = snapshot.TextBuffer;
            _loadedSnapshotVersion = snapshot.Version.VersionNumber;

            if (_anchorCache.TryGetValue(snapshot.TextBuffer, out var cached) &&
                cached.Version == snapshot.Version.VersionNumber &&
                cached.ParsedAnchors != null) {

                this.ApplyAnchors(cached.ParsedAnchors);
                return;
            }

            // Snapshot уже находится в памяти редактора, поэтому получение строк не выполняет COM-вызовов.
            var lines = snapshot.Lines.Select(line => line.GetText()).ToList();

            var anchors = State.TextEditor.AnchorParser.ParseLinesWithContextWindow(lines);
            var final = State.TextEditor.AnchorParser.InsertSeparators(anchors);

            _anchorCache[snapshot.TextBuffer] = new AnchorSnapshotCache(
                snapshot.Version.VersionNumber,
                final.Count > 0,
                final
            );

            this.ApplyAnchors(final);
        }

        private void ApplyAnchors(IReadOnlyList<State.TextEditor.AnchorPoint> anchors) {

            this.Anchors.Clear();

            // Если this.Anchors.Count > 0 - будет отображена кнопка.
            foreach (var anchor in anchors) {
                this.Anchors.Add(anchor);
            }

            bool hasAnchors = this.Anchors.Count > 0;
            this.IsAnchorToggleButtonVisible.Value = hasAnchors;
            this.IsAnchorListVisible.Value = _isAnchorListExpanded && hasAnchors;
            if (!this.IsAnchorListVisible.Value) {
                this.SelectedAnchor = null;
                return;
            }

            this.UpdateSelectionFromCaret();
        }

        private void SetActiveTextView(IWpfTextView textView) {
            if (ReferenceEquals(_activeTextView, textView)) {
                return;
            }

            this.UnsubscribeFromCaret();
            _activeTextView = textView;
            this.SubscribeToCaret();
        }

        private void SubscribeToCaret() {
            if (!_isAnchorListExpanded || _activeTextView == null) {
                return;
            }

            _activeTextView.Caret.PositionChanged -= this.OnCaretPositionChanged;
            _activeTextView.Caret.PositionChanged += this.OnCaretPositionChanged;
        }

        private void UnsubscribeFromCaret() {
            if (_activeTextView != null) {
                _activeTextView.Caret.PositionChanged -= this.OnCaretPositionChanged;
            }
        }

        private void OnCaretPositionChanged(object sender, CaretPositionChangedEventArgs e) {
            if (_isAnchorListExpanded) {
                this.UpdateSelectionFromCaret();
            }
        }

        private void UpdateSelectionFromCaret() {
            if (!_isAnchorListExpanded || _activeTextView == null || this.Anchors.Count == 0) {
                return;
            }

            int caretLineNumber = _activeTextView.Caret.Position.BufferPosition.GetContainingLine().LineNumber + 1;
            var matchingAnchor = this.Anchors
                .Where(anchor => anchor.AnchorKind != Enums.AnchorKind.Separator && anchor.LineNumber <= caretLineNumber)
                .LastOrDefault();

            _isSynchronizingCaretSelection = true;
            try {
                this.SelectedAnchor = matchingAnchor;
            }
            finally {
                _isSynchronizingCaretSelection = false;
            }
        }

        private void PrepareCollapsedState(ITextSnapshot snapshot) {
            _loadedTextBuffer = null;
            _loadedSnapshotVersion = -1;
            this.SelectedAnchor = null;
            this.Anchors.Clear();
            this.IsAnchorListVisible.Value = false;

            if (_anchorCache.TryGetValue(snapshot.TextBuffer, out var cached) && cached.Version == snapshot.Version.VersionNumber) {
                this.IsAnchorToggleButtonVisible.Value = cached.HasAnchors;
                return;
            }

            bool hasAnchors = snapshot.Lines.Any(line => line.GetText().TrimStart().StartsWith("// ░"));
            _anchorCache[snapshot.TextBuffer] = new AnchorSnapshotCache(snapshot.Version.VersionNumber, hasAnchors, null);
            this.IsAnchorToggleButtonVisible.Value = hasAnchors;
        }

        private void ResetLoadedAnchors(bool showToggleButton) {
            _activeSnapshot = null;
            _loadedTextBuffer = null;
            _loadedSnapshotVersion = -1;
            this.SelectedAnchor = null;
            this.Anchors.Clear();
            this.IsAnchorListVisible.Value = false;
            this.IsAnchorToggleButtonVisible.Value = showToggleButton;
        }

        private void NavigateToLine(int lineNumber) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (PackageServices.Dte2?.ActiveDocument?.Object("TextDocument") is EnvDTE.TextDocument textDoc) {
                var selection = textDoc.Selection;

                // Перемещаем каретку на нужную строку (на неё и останется курсор)
                selection.MoveToLineAndOffset(lineNumber, 1);

                // Для скролинга создаем точку выше — с контекстом
                int scrollLine = Math.Max(1, lineNumber - 5); // гарантируем что >= 1
                var scrollPoint = textDoc.CreateEditPoint();
                scrollPoint.MoveToLineAndOffset(scrollLine, 1);

                // Скроллим так, чтобы scrollLine оказался в самом верху
                scrollPoint.TryToShow(EnvDTE.vsPaneShowHow.vsPaneShowTop);
            }
        }
    }
}
