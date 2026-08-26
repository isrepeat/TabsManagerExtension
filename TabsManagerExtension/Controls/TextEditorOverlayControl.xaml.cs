using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Navigation;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Shell;
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
        private IAdornmentLayer? _findAdornmentLayer;
        private ITextBuffer? _loadedTextBuffer;
        private int _loadedSnapshotVersion = -1;
        private bool _isAnchorListExpanded;
        private bool _isSynchronizingCaretSelection;
        private bool _isEditorFrameActive = true;
        private bool _keepFindCommandsWhenEditorInactive;
        private readonly Dictionary<ITextBuffer, AnchorSnapshotCache> _anchorCache = new();

        private sealed class AnchorSnapshotCache {
            public int Version { get; }
            public int PatternsRevision { get; }
            public bool HasAnchors { get; }
            public IReadOnlyList<State.TextEditor.AnchorPoint>? ParsedAnchors { get; }

            public AnchorSnapshotCache(
                int version,
                int patternsRevision,
                bool hasAnchors,
                IReadOnlyList<State.TextEditor.AnchorPoint>? parsedAnchors
                ) {
                this.Version = version;
                this.PatternsRevision = patternsRevision;
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
        public ICommand OnFindNextCommand { get; }
        public ICommand OnFindPreviousCommand { get; }
        public ICommand OnFindInSolutionCommand { get; }

        public TextEditorOverlayControl() {
            this.InitializeComponent();
            this.Loaded += this.OnLoaded;
            this.Unloaded += this.OnUnloaded;
            this.SizeChanged += this.OnOverlaySizeChanged;
            this.DataContext = this;

            this.OnToggleAnchorListCommand = new Helpers.RelayCommand<object>(this.OnToggleAnchorList);
            this.OnFindNextCommand = new Helpers.RelayCommand(this.FindNext);
            this.OnFindPreviousCommand = new Helpers.RelayCommand(this.FindPrevious);
            this.OnFindInSolutionCommand = new Helpers.RelayCommand(this.FindInSolution);
        }

        private void OnLoaded(object sender, RoutedEventArgs e) {
            //VsShell.TextEditor.Services.TextEditorCommandFilterService.Instance.AddTrackedInputElement(this);

            // IsHitTestVisible могут быть унаследованы от родителя (например, AdornerLayer),
            // поэтому значения из XAML не применяются гарантированно — устанавливаем явно в OnLoaded.
            this.IsHitTestVisible = true;
            this.UpdateAnchorListMaxHeight();
            this.SubscribeToFindAdornmentChanges();
            this.UpdateAnchorContainerVisibility();
            Settings.TabsManagerSettingsService.AnchorPatternsChanged += this.OnAnchorPatternsChanged;

        }

        private void OnOverlaySizeChanged(object sender, SizeChangedEventArgs e) {
            this.UpdateAnchorListMaxHeight();
        }

        private void UpdateAnchorListMaxHeight() {
            // Оставляем место для верхнего отступа, кнопки, промежутка и нижней границы редактора.
            const double reservedHeight = 30 + 42 + 5 + 10;
            this.AnchorListBox.MaxHeight = Math.Max(0, this.ActualHeight - reservedHeight);
        }

        private void OnUnloaded(object sender, RoutedEventArgs e) {
            Settings.TabsManagerSettingsService.AnchorPatternsChanged -= this.OnAnchorPatternsChanged;
            this.UnsubscribeFromFindAdornmentChanges();
            //VsShell.TextEditor.Services.TextEditorCommandFilterService.Instance.RemoveTrackedInputElement(this);
            // Контрол временно выгружается при переносе adorner между document frame.
            // Не очищаем состояние и кэш: новый snapshot будет применён сразу после повторного подключения.
        }

        private void OnAnchorPatternsChanged() {
            this.Dispatcher.InvokeAsync(() => {
                _anchorCache.Clear();
                if (_activeSnapshot == null) {
                    return;
                }

                if (_isAnchorListExpanded) {
                    this.LoadAnchors(_activeSnapshot);
                }
                else {
                    this.PrepareCollapsedState(_activeSnapshot);
                }
            });
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
            this.UnsubscribeFromFindAdornmentChanges();
            _activeTextView = null;
            _findAdornmentLayer = null;
            this.ResetLoadedAnchors(showToggleButton: false);
        }

        public void OnEditorFrameActivityChanged(bool isActive, bool keepFindCommandsWhenInactive) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _isEditorFrameActive = isActive;
            _keepFindCommandsWhenEditorInactive = keepFindCommandsWhenInactive;
            this.UpdateAnchorContainerVisibility();
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
                cached.PatternsRevision == Settings.TabsManagerSettingsService.AnchorPatternsRevision &&
                cached.ParsedAnchors != null) {

                this.ApplyAnchors(cached.ParsedAnchors);
                return;
            }

            // Snapshot уже находится в памяти редактора, поэтому получение строк не выполняет COM-вызовов.
            var lines = snapshot.Lines.Select(line => line.GetText()).ToList();

            var anchors = State.TextEditor.AnchorParser.ParseLinesWithContextWindow(
                lines,
                Settings.TabsManagerSettingsService.AnchorSectionPattern,
                Settings.TabsManagerSettingsService.AnchorSubsectionPattern
            );
            var final = State.TextEditor.AnchorParser.InsertSeparators(anchors);

            _anchorCache[snapshot.TextBuffer] = new AnchorSnapshotCache(
                snapshot.Version.VersionNumber,
                Settings.TabsManagerSettingsService.AnchorPatternsRevision,
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
                this.SubscribeToFindAdornmentChanges();
                this.UpdateAnchorContainerVisibility();
                return;
            }

            this.UnsubscribeFromCaret();
            this.UnsubscribeFromFindAdornmentChanges();
            _activeTextView = textView;
            _findAdornmentLayer = textView.GetAdornmentLayer("FindUIAdornmentLayer");
            this.SubscribeToFindAdornmentChanges();
            this.UpdateAnchorContainerVisibility();
            this.SubscribeToCaret();
        }

        private void SubscribeToFindAdornmentChanges() {
            if (_activeTextView == null || !this.IsLoaded) {
                return;
            }

            _activeTextView.VisualElement.LayoutUpdated -= this.OnTextViewLayoutUpdated;
            _activeTextView.VisualElement.LayoutUpdated += this.OnTextViewLayoutUpdated;
        }

        private void UnsubscribeFromFindAdornmentChanges() {
            if (_activeTextView != null) {
                _activeTextView.VisualElement.LayoutUpdated -= this.OnTextViewLayoutUpdated;
            }
        }

        private void OnTextViewLayoutUpdated(object sender, EventArgs e) {
            this.UpdateAnchorContainerVisibility();
        }

        private void UpdateAnchorContainerVisibility() {
            // Штатная Quick Find занимает тот же верхний правый угол, поэтому якоря и команды не показываются одновременно.
            bool isFindVisible = _findAdornmentLayer?.IsEmpty == false;
            bool showFindCommands = isFindVisible &&
                (_isEditorFrameActive || _keepFindCommandsWhenEditorInactive);
            this.Visibility = _isEditorFrameActive || showFindCommands
                ? Visibility.Visible
                : Visibility.Collapsed;
            this.AnchorContainer.Visibility = _isEditorFrameActive && !isFindVisible
                ? Visibility.Visible
                : Visibility.Collapsed;
            this.FindCommandsContainer.Visibility = showFindCommands
                ? Visibility.Visible
                : Visibility.Collapsed;

            if (showFindCommands) {
                this.UpdateFindCommandsPosition();
            }
        }

        private void UpdateFindCommandsPosition() {
            const double fallbackTop = 135;
            const double verticalGap = 4;
            double? top = null;

            if (_findAdornmentLayer != null) {
                foreach (var element in _findAdornmentLayer.Elements) {
                    var adornment = element.Adornment as FrameworkElement;
                    if (adornment?.IsVisible != true || adornment.ActualHeight <= 0) {
                        continue;
                    }

                    try {
                        var topLeft = adornment.TranslatePoint(new Point(0, 0), this);
                        top = Math.Max(verticalGap, topLeft.Y + adornment.ActualHeight + verticalGap);
                        break;
                    }
                    catch (InvalidOperationException) {
                        // В разных версиях VS Find UI может находиться в отдельной visual-ветке.
                    }
                }
            }

            var margin = new Thickness(0, top ?? fallbackTop, 30, 0);
            if (this.FindCommandsContainer.Margin != margin) {
                this.FindCommandsContainer.Margin = margin;
            }
        }

        private void FindNext() {
            ThreadHelper.ThrowIfNotOnUIThread();
            PackageServices.Dte2.ExecuteCommand("Edit.FindNext");
        }

        private void FindPrevious() {
            ThreadHelper.ThrowIfNotOnUIThread();
            PackageServices.Dte2.ExecuteCommand("Edit.FindPrevious");
        }

        private void FindInSolution() {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Запускаем поиск через уже открытую Quick Find: так VS использует современное окно результатов, а не legacy Find Results 1.
            if (!this.TryGetSolutionSearchScope(out var scope, out var findManager)) {
                Helpers.Diagnostic.Logger.LogWarning("[TextEditorOverlay] Quick Find solution scope was not found.");
                return;
            }

            findManager.GetType().GetProperty("CurrentScope")?.SetValue(findManager, scope);

            var searchScopeInterface = scope.GetType().GetInterfaces()
                .First(type => type.FullName == "Microsoft.VisualStudio.Find.ISearchScope");
            var findAllMethod = searchScopeInterface.GetMethod("DoFindAllAsync");
            if (findAllMethod?.Invoke(scope, new object[] { CancellationToken.None }) is Task findTask) {
                _ = findTask.ContinueWith(
                    task => Helpers.Diagnostic.Logger.LogError($"[TextEditorOverlay] Find in solution failed: {task.Exception}"),
                    CancellationToken.None,
                    TaskContinuationOptions.OnlyOnFaulted,
                    TaskScheduler.Default
                );
            }
        }

        private bool TryGetSolutionSearchScope(out object scope, out object findManager) {
            scope = null!;
            findManager = null!;

            if (_findAdornmentLayer == null) {
                return false;
            }

            foreach (var element in _findAdornmentLayer.Elements) {
                // Типы Quick Find не входят в публичный API, поэтому извлекаем менеджер из visual-дерева без жёсткой зависимости от версии VS.
                var candidateManager = TryGetQuickFindManager(element.Adornment);
                if (candidateManager?.GetType().GetProperty("ScopesCollection")?.GetValue(candidateManager) is not System.Collections.IEnumerable scopes) {
                    continue;
                }

                foreach (var item in scopes) {
                    var candidateScope = item == null ? null : TryUnwrapSearchScope(item);
                    if (candidateScope == null || !IsSolutionSearchScope(candidateScope)) {
                        continue;
                    }

                    scope = candidateScope;
                    findManager = candidateManager;
                    return true;
                }
            }

            return false;
        }

        private static object? TryGetQuickFindManager(UIElement findAdornment) {
            if (findAdornment is FrameworkElement frameworkElement && IsQuickFindManager(frameworkElement.DataContext)) {
                return frameworkElement.DataContext;
            }

            var adornmentType = findAdornment.GetType();
            foreach (var property in adornmentType.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)) {
                if (property.GetIndexParameters().Length != 0 || property.PropertyType.Name.IndexOf("QuickFindManager", StringComparison.Ordinal) < 0) {
                    continue;
                }

                try {
                    var value = property.GetValue(findAdornment);
                    if (IsQuickFindManager(value)) {
                        return value;
                    }
                }
                catch (System.Reflection.TargetInvocationException) {
                }
            }

            foreach (var field in adornmentType.GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)) {
                var value = field.GetValue(findAdornment);
                if (IsQuickFindManager(value)) {
                    return value;
                }
            }

            return null;
        }

        private static bool IsQuickFindManager(object? value) {
            return value?.GetType().Name.IndexOf("QuickFindManager", StringComparison.Ordinal) >= 0;
        }

        private static object? TryUnwrapSearchScope(object item) {
            object? current = item;
            for (int depth = 0; depth < 3 && current != null; depth++) {
                if (current.GetType().GetInterfaces().Any(type => type.FullName == "Microsoft.VisualStudio.Find.ISearchScope")) {
                    return current;
                }

                current = current.GetType().GetProperty("Item")?.GetValue(current);
            }

            return null;
        }

        private static bool IsSolutionSearchScope(object scope) {
            if (scope.GetType().Name.IndexOf("SolutionScope", StringComparison.Ordinal) >= 0) {
                return true;
            }

            var searchScopeInterface = scope.GetType().GetInterfaces()
                .First(type => type.FullName == "Microsoft.VisualStudio.Find.ISearchScope");
            string? displayName = searchScopeInterface.GetProperty("DisplayName")?.GetValue(scope) as string;
            return string.Equals(displayName, "Entire solution", StringComparison.OrdinalIgnoreCase);
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

            if (_anchorCache.TryGetValue(snapshot.TextBuffer, out var cached) &&
                cached.Version == snapshot.Version.VersionNumber &&
                cached.PatternsRevision == Settings.TabsManagerSettingsService.AnchorPatternsRevision) {

                this.IsAnchorToggleButtonVisible.Value = cached.HasAnchors;
                return;
            }

            var lines = snapshot.Lines.Select(line => line.GetText()).ToList();
            bool hasAnchors = State.TextEditor.AnchorParser.ContainsAnchor(
                lines,
                Settings.TabsManagerSettingsService.AnchorSectionPattern,
                Settings.TabsManagerSettingsService.AnchorSubsectionPattern
            );
            _anchorCache[snapshot.TextBuffer] = new AnchorSnapshotCache(
                snapshot.Version.VersionNumber,
                Settings.TabsManagerSettingsService.AnchorPatternsRevision,
                hasAnchors,
                null
            );
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
