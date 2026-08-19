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
        private ITextBuffer? _loadedTextBuffer;
        private int _loadedSnapshotVersion = -1;
        public State.TextEditor.AnchorPoint? SelectedAnchor {
            get => _selectedAnchor;
            set {
                if (_selectedAnchor != value) {
                    _selectedAnchor = value;
                    this.OnPropertyChanged();
                    if (value != null) {
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

            this.LoadAnchorsFromActiveDocument();
        }

        private void OnUnloaded(object sender, RoutedEventArgs e) {
            //VsShell.TextEditor.Services.TextEditorCommandFilterService.Instance.RemoveTrackedInputElement(this);
            this.Anchors.Clear();
        }


        private void OnToggleAnchorList(object parameter) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnToggleAnchorList()");
            this.IsAnchorListVisible.Value = !this.IsAnchorListVisible.Value;

            if (this.IsAnchorListVisible.Value == false) {
                this.SelectedAnchor = null;
            }
        }


        public void LoadAnchorsFromActiveDocument() {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Берём текст напрямую из буфера редактора, чтобы не читать документ построчно через медленный DTE/COM API.
            var viewHost = VsShell.TextEditor.TextEditorControlHelper.TryGetActiveViewHost();
            var snapshot = viewHost?.TextView.TextSnapshot;
            if (snapshot == null) {
                return;
            }

            // При возврате из Output активируется тот же snapshot — повторный разбор в этом случае не нужен.
            if (ReferenceEquals(_loadedTextBuffer, snapshot.TextBuffer) && _loadedSnapshotVersion == snapshot.Version.VersionNumber) {
                return;
            }

            _loadedTextBuffer = snapshot.TextBuffer;
            _loadedSnapshotVersion = snapshot.Version.VersionNumber;

            // Snapshot уже находится в памяти редактора, поэтому получение строк не выполняет COM-вызовов.
            var lines = snapshot.Lines.Select(line => line.GetText()).ToList();

            var anchors = State.TextEditor.AnchorParser.ParseLinesWithContextWindow(lines);
            var final = State.TextEditor.AnchorParser.InsertSeparators(anchors);

            this.Anchors.Clear();

            // Если this.Anchors.Count > 0 - будет отображена кнопка.
            foreach (var anchor in final) {
                this.Anchors.Add(anchor);
            }
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
