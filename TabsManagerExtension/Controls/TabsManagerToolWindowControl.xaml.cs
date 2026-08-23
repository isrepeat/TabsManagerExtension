using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.ComponentModel.Design;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using Helpers.Ex;
using TabsManagerExtension.State.Document;

// TODO: Вынеси UndoRedo логику табов в отдельный класс

namespace TabsManagerExtension.Controls {
    public partial class TabsManagerToolWindowControl : Helpers.BaseUserControl {

        // Properties:
        private Helpers.Collections.SortedObservableCollection<TabItemsGroupBase> _sortedTabItemsGroups;
        public Helpers.Collections.SortedObservableCollection<TabItemsGroupBase> SortedTabItemsGroups {
            get => _sortedTabItemsGroups;
            private set {
                if (_sortedTabItemsGroups != value) {
                    _sortedTabItemsGroups = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<Helpers.IMenuItem> _contextMenuItems;
        public ObservableCollection<Helpers.IMenuItem> ContextMenuItems {
            get => _contextMenuItems;
            private set {
                if (_contextMenuItems != value) {
                    _contextMenuItems = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<Helpers.IMenuItem> _virtualMenuItems;
        public ObservableCollection<Helpers.IMenuItem> VirtualMenuItems {
            get => _virtualMenuItems;
            private set {
                if (_virtualMenuItems != value) {
                    _virtualMenuItems = value;
                    OnPropertyChanged();
                }
            }
        }

        private ObservableCollection<Helpers.IMenuItem> _toolbarMenuItems;
        public ObservableCollection<Helpers.IMenuItem> ToolbarMenuItems {
            get => _toolbarMenuItems;
            private set {
                if (_toolbarMenuItems != value) {
                    _toolbarMenuItems = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<State.BackgroundOperationStatus> BackgroundOperations =>
            Services.ExtensionStatusService.Instance.Operations;

        private bool _isIncludeGraphReady;
        public bool IsIncludeGraphReady {
            get => _isIncludeGraphReady;
            private set {
                if (_isIncludeGraphReady != value) {
                    _isIncludeGraphReady = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _scaleFactorUI = 1.0;
        public double ScaleFactorUI {
            get => _scaleFactorUI;
            set {
                if (_scaleFactorUI != value) {
                    _scaleFactorUI = value;
                    OnPropertyChanged();
                    this.ApplyScaleUI();
                }
            }
        }

        private double _scaleFactorTabsCompactness = 1.0;
        public double ScaleFactorTabsCompactness {
            get => _scaleFactorTabsCompactness;
            set {
                if (_scaleFactorTabsCompactness != value) {
                    _scaleFactorTabsCompactness = value;
                    OnPropertyChanged();
                    this.ApplyScaleTabsCompactness();
                    Configuration.TabsManagerConfigurationService.SetTabsScaleFactor(value);
                }
            }
        }

        private bool _isTabEditMode;
        public bool IsTabEditMode {
            get => _isTabEditMode;
            set {
                if (_isTabEditMode != value) {
                    _isTabEditMode = value;
                    OnPropertyChanged();

                    if (value) {
                        this.UpdateEditModeInputRedirect();
                        this.Dispatcher.BeginInvoke(new Action(() => _keyboardTabNavigationExtension?.InitializeFocus()));
                    }
                    else {
                        // Выход из режима навигации убирает только пунктирный фокус.
                        // Сформированный пользователем мультивыбор должен сохраниться.
                        this.UpdateEditModeInputRedirect();
                        _keyboardTabNavigationExtension?.ClearFocus();
                    }
                }
            }
        }

        private bool _isMultipleTabSelection;
        public bool IsMultipleTabSelection {
            get => _isMultipleTabSelection;
            private set {
                if (_isMultipleTabSelection != value) {
                    _isMultipleTabSelection = value;
                    OnPropertyChanged();
                }
            }
        }


        // Internal:
        private EnvDTE.WindowEvents _windowEvents;
        private EnvDTE.DocumentEvents _documentEvents;
        private EnvDTE.SolutionEvents _solutionEvents;

        private DispatcherTimer? _tabsManagerStateTimer;
        private FileSystemWatcher _fileWatcher;
        private bool _isRestoringToolWindows;
        // Во время Ctrl+Z активация открываемых документов не должна учитывать физически
        // удерживаемый Ctrl и добавлять вкладки к существующему мультивыбору.
        private bool _isRestoringClosedTabs;
        private bool _isSolutionClosing;
        private string? _openDocumentsLoadedForSolution;
        // Одна запись стека соответствует одному пользовательскому закрытию. Поэтому вкладки,
        // закрытые через мультивыбор, восстанавливаются одним нажатием Ctrl+Z.
        private readonly Stack<ClosedTabsOperation> _closedTabsHistory = new();
        // Закрытие через CloseTabItems сначала записывает всю операцию. Ключи не дают событиям
        // DocumentClosing/WindowClosing повторно создать по одной записи для каждой вкладки.
        private readonly HashSet<string> _tabsBeingClosed = new(StringComparer.OrdinalIgnoreCase);
        private const int ClosedTabsHistoryCapacity = 50;

        private Helpers.Collections.GroupsSelectionCoordinator<TabItemsGroupBase, TabItemBase> _tabItemsSelectionCoordinator;
        private Navigation.TabNavigationController _tabNavigationController;
        private Navigation.KeyboardTabNavigationExtension _keyboardTabNavigationExtension;
        private VsShell.TextEditor.Overlay.TextEditorOverlayController _textEditorOverlayController;
        private readonly HashSet<TabItemControl> _tabItemControls = new();
        // Текущий UX: только обычный ЛКМ меняет активный документ; Ctrl/Shift/Space
        // управляют selection, не перемещая фиолетовую рамку активного VS-фрейма.
        private const Navigation.TabSelectionActivationPolicy TabSelectionActivationPolicy =
            Navigation.TabSelectionActivationPolicy.ActivateOnlyOnUnmodifiedPointerSelection;
        private static readonly HashSet<string> RenamableTabExtensions = new(StringComparer.OrdinalIgnoreCase) {
            ".c",
            ".cpp",
            ".h",
            ".cxx",
            ".hxx"
        };

        public ICommand OnPinTabItemCommand { get; }
        public ICommand OnUnpinTabItemCommand { get; }
        public ICommand OnCloseTabItemCommand { get; }
        public ICommand OnKeepOpenedTabItemCommand { get; }
        public ICommand OnTabItemContextMenuOpenCommand { get; }
        public ICommand OnTabItemContextMenuClosedCommand { get; }
        public ICommand OnTabItemVirtualMenuOpenCommand { get; }
        public ICommand OnTabItemVirtualMenuClosedCommand { get; }
        public ICommand OnProjectContextIncludersOpenCommand { get; }
        public ICommand OnProjectContextMenuItemMouseEnterCommand { get; }
        public ICommand OnToolbarMenuShowCommand { get; }
        public ICommand OnToolbarMenuOpenCommand { get; }
        public ICommand OnToolbarMenuClosedCommand { get; }

        public TabsManagerToolWindowControl() {
            this.InitializeComponent();
            this.ScaleFactorTabsCompactness = Configuration.TabsManagerConfigurationService.TabsScaleFactor;
            this.Loaded += this.OnLoaded;
            this.Unloaded += this.OnUnloaded;
            this.PreviewMouseDown += this.OnPreviewMouseDown;
            this.PreviewKeyDown += this.OnPreviewKeyDown;
            this.KeyDown += this.OnPreviewKeyDown;

            this.OnPinTabItemCommand = new Helpers.RelayCommand<object>(this.OnPinTabItem);
            this.OnUnpinTabItemCommand = new Helpers.RelayCommand<object>(this.OnUnpinTabItem);
            this.OnCloseTabItemCommand = new Helpers.RelayCommand<object>(this.OnCloseTabItem);
            this.OnKeepOpenedTabItemCommand = new Helpers.RelayCommand<object>(this.OnKeepOpenedTabItem);

            this.OnTabItemContextMenuOpenCommand = new Helpers.RelayCommand<object>(this.OnTabItemContextMenuOpen);
            this.OnTabItemContextMenuClosedCommand = new Helpers.RelayCommand<object>(this.OnTabItemContextMenuClosed);
            
            this.OnTabItemVirtualMenuOpenCommand = new Helpers.RelayCommand<object>(this.OnTabItemVirtualMenuOpen);
            this.OnTabItemVirtualMenuClosedCommand = new Helpers.RelayCommand<object>(this.OnTabItemVirtualMenuClosed);
            this.OnProjectContextIncludersOpenCommand = new Helpers.RelayCommand<object>(this.OnProjectContextIncludersOpen);
            this.OnProjectContextMenuItemMouseEnterCommand = new Helpers.RelayCommand<object>(this.OnProjectContextMenuItemMouseEnter);
            this.OnToolbarMenuShowCommand = new Helpers.RelayCommand<object>(this.OnToolbarMenuShow);
            this.OnToolbarMenuOpenCommand = new Helpers.RelayCommand<object>(this.OnToolbarMenuOpen);
            this.OnToolbarMenuClosedCommand = new Helpers.RelayCommand<object>(this.OnToolbarMenuClosed);

            // Коллекция должна существовать до первого показа: в отличие от старой реализации,
            // UpdateVirtualMenuItems обновляет её на месте, а не заменяет новым экземпляром.
            this.VirtualMenuItems = new ObservableCollection<Helpers.IMenuItem>();

            base.DataContext = this;
        }


        //
        // ░ Self handlers
        // ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
        //
        private void OnToolbarMenuShow(object parameter) {
            if (parameter is not Button button) {
                return;
            }

            this.ToolbarMenu.ShowMenu(
                button,
                System.Windows.Controls.Primitives.PlacementMode.Bottom,
                false,
                placementTarget: button
            );
        }

        private void OnToolbarMenuOpen(object parameter) {
            // Пункты создаются при каждом открытии, как у контекстного меню таба:
            // актуальные команды и их доступность определяются непосредственно перед показом popup.
            var unavailableCommand = new Helpers.RelayCommand(() => { }, () => false);
            this.ToolbarMenuItems = new ObservableCollection<Helpers.IMenuItem> {
                new Helpers.MenuItemCommand {
                    Header = "Save tab layout (coming soon)",
                    Command = unavailableCommand
                },
                new Helpers.MenuItemCommand {
                    Header = "Restore tab layout (coming soon)",
                    Command = unavailableCommand
                },
                new Helpers.MenuItemSeparator(),
                new Helpers.MenuItemCommand {
                    Header = "Options…",
                    Command = new Helpers.RelayCommand(TabsManagerExtensionPackage.ShowOptions)
                }
            };
        }

        private void OnToolbarMenuClosed(object parameter) {
            // Освобождаем созданные для текущего открытия пункты и связанные с ними команды.
            this.ToolbarMenuItems = new ObservableCollection<Helpers.IMenuItem>();
        }


        private void OnLoaded(object sender, RoutedEventArgs e) {
            Services.ExtensionServices.BeginUsage();
            Configuration.TabsManagerConfigurationService.TabsScaleFactorChanged += this.OnTabsScaleFactorChanged;
            VsShell.TextEditor.Services.TextEditorInputCommandFilterService.Instance.AddTrackedInputElement(this);
            this.UpdateEditModeInputRedirect();
            Services.ExtensionStatusService.Instance.FeatureReadinessChanged += this.OnFeatureReadinessChanged;
            this.IsIncludeGraphReady = Services.ExtensionStatusService.Instance.IsFeatureReady(
                Services.ExtensionStatusService.IncludeGraphFeature
            );

            this.InitializeDTE();
            this.InitializeFileWatcher();
            this.InitializeVsShellTrackers();
            this.InitializeTabItemsSelectionCoordinator();
            this.InitializeBackgroundRoutine();
            this.ApplyScaleTabsCompactness();

            var hierarchyAnalyzer = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance;
            hierarchyAnalyzer.InitialAnalysisCompleted.Add(this.OnInitialHierarchyAnalysisCompleted);
            hierarchyAnalyzer.InitialAnalysisCompleted.InvokeForLastHandlerIfTriggered();
        }

        private void OnUnloaded(object sender, RoutedEventArgs e) {
            Configuration.TabsManagerConfigurationService.TabsScaleFactorChanged -= this.OnTabsScaleFactorChanged;
            VsShell.TextEditor.Services.TextEditorInputCommandFilterService.Instance.SetForcedInputTarget(null);
            VsShell.TextEditor.Services.TextEditorInputCommandFilterService.Instance.RemoveTrackedInputElement(this);
            VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance.InitialAnalysisCompleted.Remove(this.OnInitialHierarchyAnalysisCompleted);
            Services.ExtensionStatusService.Instance.FeatureReadinessChanged -= this.OnFeatureReadinessChanged;
            this.SaveOpenToolWindows();
            this.UninitializeTabItemsSelectionCoordinator();
            this.UninitializeVsShellTrackers();
            this.UninitializeFileWatcher();
            this.UninitializeBackgroundRoutine();
            this.UninitializeDTE();

            Services.ExtensionServices.EndUsage();
        }

        private void OnTabsScaleFactorChanged(double value) {
            this.Dispatcher.InvokeAsync(() => this.ScaleFactorTabsCompactness = value);
        }

        private void OnInitialHierarchyAnalysisCompleted(string solutionName) {
            if (string.Equals(
                _openDocumentsLoadedForSolution,
                solutionName,
                StringComparison.OrdinalIgnoreCase)) {

                return;
            }

            _openDocumentsLoadedForSolution = solutionName;
            this.LoadOpenDocuments();
        }

        private void OnFeatureReadinessChanged(string feature, bool isReady) {
            if (feature == Services.ExtensionStatusService.IncludeGraphFeature) {
                this.IsIncludeGraphReady = isReady;
            }
        }

        private void OnPreviewMouseDown(object sender, MouseButtonEventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // У DocumentContainer намеренно нет Background: пустая область не должна сама становиться
            // WPF-целью ввода. Поэтому определяем попадание по координатам на корневом PreviewMouseDown.
            var pointerPosition = e.GetPosition(this.DocumentContainer);
            bool isInsideDocumentContainer =
                pointerPosition.X >= 0 &&
                pointerPosition.Y >= 0 &&
                pointerPosition.X <= this.DocumentContainer.ActualWidth &&
                pointerPosition.Y <= this.DocumentContainer.ActualHeight;

            var originalSource = e.OriginalSource as DependencyObject;
            bool isTabInteraction = originalSource is TabItemControl ||
                originalSource != null && Helpers.VisualTree.FindParentByType<TabItemControl>(originalSource) != null;

            bool isButtonInteraction = originalSource is System.Windows.Controls.Primitives.ButtonBase ||
                originalSource != null && Helpers.VisualTree.FindParentByType<System.Windows.Controls.Primitives.ButtonBase>(originalSource) != null;

            // Клик по вкладке или кнопке имеет собственную семантику и не считается кликом по пустой области.
            if (isInsideDocumentContainer && !isTabInteraction && !isButtonInteraction) {
                // Сбрасываем мультивыбор явно: повторная активация уже открытого документа
                // может не породить событие DTE после команд контекстного меню.
                var activeFrameTabItem = this.SortedTabItemsGroups
                    .SelectMany(group => group.Items)
                    .FirstOrDefault(tabItem => tabItem.Metadata.GetFlag("IsFrameActive"));

                // Фиолетовая рамка отражает фактический VS-фрейм и имеет приоритет над PrimarySelection,
                // которая при мультивыборе может указывать на совсем другую вкладку.
                var retainedTabItem = activeFrameTabItem ?? _tabItemsSelectionCoordinator.PrimarySelection?.Item;

                if (retainedTabItem != null) {
                    _tabNavigationController.SetSelectionWithoutActivation(retainedTabItem, true, ModifierKeys.None);
                }

                if (this.IsTabEditMode) {
                    // В режиме навигации пустая область активирует панель: дальнейшие стрелки
                    // должны продолжать перемещать пунктирный навигационный фокус по вкладкам.
                    if (retainedTabItem != null) {
                        _keyboardTabNavigationExtension.FocusItem(retainedTabItem);
                    }

                    // FocusedItem мог не измениться, поэтому восстанавливаем input target явно.
                    this.FocusEditModeInputTarget();
                    Helpers.GlobalFlags.SetFlag("TextEditorFrameFocused", false);
                }
                else {
                    // В обычном режиме панель не удерживает фокус — возвращаем его активному
                    // документу или tool window, сохранив при этом одиночное выделение вкладки.
                    if (retainedTabItem is IActivatableTab activatableTab) {
                        activatableTab.Activate(); // Инициирует фокус редактора (или окна на его месте).
                    }

                    Helpers.GlobalFlags.SetFlag("TextEditorFrameFocused", true);
                }

                e.Handled = true;
            }
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e) {
            if (!this.IsTabEditMode) {
                return;
            }

            // Пока открыт inline rename, клавиши принадлежат TextBox. В частности, корневой
            // PreviewKeyDown не должен превратить Enter в активацию вкладки раньше поля ввода.
            if (Keyboard.FocusedElement is DependencyObject focusedElement) {
                var focusedTabControl = Helpers.VisualTree.FindParentByType<TabItemControl>(focusedElement);
                if (focusedTabControl?.IsRenaming == true) {
                    return;
                }
            }

            // Корневой обработчик не зависит от существования визуального TabItemControl:
            // Ctrl+Z остаётся доступен и после удаления последней вкладки.
            if (this.HandleTabEditKey(e.Key, Keyboard.Modifiers)) {
                e.Handled = true;
            }
        }

        private void UpdateEditModeInputRedirect() {
            if (!this.IsLoaded) {
                return;
            }

            // PART_TabListHost не становится отдельным command target. Пока активен edit mode,
            // перехватываем навигационные команды в IVsTextView и направляем их в этот контрол.
            VsShell.TextEditor.Services.TextEditorInputCommandFilterService.Instance.SetForcedInputTarget(
                this.IsTabEditMode ? this : null,
                this.IsTabEditMode ? this.CanHandleRedirectedNavigationKey : null
            );
        }


        //
        // ░ Initialization
        // ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 
        // 
        // ░ DTE
        //
        private void InitializeDTE() {
            ThreadHelper.ThrowIfNotOnUIThread();

            _documentEvents = PackageServices.Dte2.Events.DocumentEvents;
            _documentEvents.DocumentOpened += OnDocumentOpened;
            _documentEvents.DocumentSaved += OnDocumentSaved;
            _documentEvents.DocumentClosing += OnDocumentClosing;

            _windowEvents = PackageServices.Dte2.Events.WindowEvents;
            _windowEvents.WindowActivated += OnWindowActivated;
            _windowEvents.WindowClosing += OnWindowClosing;

            _solutionEvents = PackageServices.Dte2.Events.SolutionEvents;
            _solutionEvents.BeforeClosing += OnSolutionClosing;
        }

        private void UninitializeDTE() {
            ThreadHelper.ThrowIfNotOnUIThread();

            _solutionEvents.BeforeClosing -= OnSolutionClosing;

            _windowEvents.WindowClosing -= OnWindowClosing;
            _windowEvents.WindowActivated -= OnWindowActivated;

            _documentEvents.DocumentClosing -= OnDocumentClosing;
            _documentEvents.DocumentSaved -= OnDocumentSaved;
            _documentEvents.DocumentOpened -= OnDocumentOpened;
        }


        //
        // ░ FileWatcher 
        //
        private void InitializeFileWatcher() {
            ThreadHelper.ThrowIfNotOnUIThread();

            var solution = PackageServices.Dte2.Solution;
            var solutionDir = string.IsNullOrEmpty(solution.FullName) 
                ? null 
                : Path.GetDirectoryName(solution.FullName);

            if (string.IsNullOrEmpty(solutionDir)) {
                return;
            }

            _fileWatcher = new FileSystemWatcher {
                Path = solutionDir,
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size,
                IncludeSubdirectories = true,
                Filter = "*.*"
            };

            _fileWatcher.Changed += OnFileChanged;
            _fileWatcher.Renamed += OnFileRenamed;
            _fileWatcher.Deleted += OnFileDeleted;
            _fileWatcher.EnableRaisingEvents = true;
        }

        private void UninitializeFileWatcher() {
            if (_fileWatcher != null) {
                try {
                    _fileWatcher.EnableRaisingEvents = false;

                    // Удаляем обработчики событий, чтобы отложенные события не вызывались
                    _fileWatcher.Changed -= this.OnFileChanged;
                    _fileWatcher.Renamed -= this.OnFileRenamed;
                    _fileWatcher.Deleted -= this.OnFileDeleted;

                    var watcherToDispose = _fileWatcher;
                    _fileWatcher = null;

                    // DispatcherPriority.ApplicationIdle — ждет, пока текущий UI-цикл и все запланированные задачи завершатся.
                    // Таким образом Dispose() вызывается после завершения Run(...) внутри событий FileSystemWatcher
                    VsixThreadHelper.RunOnUiThread(Dispatcher, () => {
                        try {
                            // Копируем ссылку на _fileWatcher чтобы продлить жизнь, т.к. _fileWatcher уже может быть удален
                            // во время исполнения лямбды, но его ресурсы так и не будут освобождены.
                            watcherToDispose.Dispose();
                        }
                        catch (Exception ex) {
                            Helpers.Diagnostic.Logger.LogError($"Delayed dispose of FileSystemWatcher failed: {ex}");
                        }
                    }, DispatcherPriority.ApplicationIdle);
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogError($"Error while scheduling FileSystemWatcher disposal: {ex}");
                }
            }
        }


        //
        // ░ VsShellTrackers 
        //
        private void InitializeVsShellTrackers() {
            VsShell.Document.Services.VsDocumentActivationTrackerService.Instance.OnDocumentActivated += this.OnDocumentActivatedExternally;
            VsShell.Solution.Services.VsWindowFrameActivationTrackerService.Instance.VsWindowFrameActivated += this.OnVsWindowFrameActivated;
            VsShell.TextEditor.Services.TextEditorFileNavigationCommandFilterService.Instance.OnNavigatedToDocument += this.OnTextEditorNavigatedToDocument;
        }
        private void UninitializeVsShellTrackers() {
            VsShell.TextEditor.Services.TextEditorFileNavigationCommandFilterService.Instance.OnNavigatedToDocument -= this.OnTextEditorNavigatedToDocument;
            VsShell.Solution.Services.VsWindowFrameActivationTrackerService.Instance.VsWindowFrameActivated -= this.OnVsWindowFrameActivated;
            VsShell.Document.Services.VsDocumentActivationTrackerService.Instance.OnDocumentActivated -= this.OnDocumentActivatedExternally;
        }


        //
        // ░ TabItemsSelectionCoordinator 
        //
        private void InitializeTabItemsSelectionCoordinator() {
            var defaultTabItemsGroupComparer = Comparer<TabItemsGroupBase>.Create((a, b) => string.Compare(a.GroupName, b.GroupName, StringComparison.OrdinalIgnoreCase));
            var priorityGroups = new List<Helpers.Collections.PriorityGroup<TabItemsGroupBase>> {
                new Helpers.Collections.PriorityGroup<TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top,
                    InsertMode = Helpers.Collections.ItemInsertMode.SingleWithReplaceExisting,
                    Predicate = g => g is TabItemsPreviewGroup,
                    Comparer = defaultTabItemsGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top + 1,
                    InsertMode = Helpers.Collections.ItemInsertMode.Single,
                    Predicate = g => g is SeparatorTabItemsGroup separator && separator.Key == "Preview-Pinned",
                    Comparer = defaultTabItemsGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top + 2,
                    Predicate = g => g is TabItemsPinnedGroup,
                    Comparer = defaultTabItemsGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top + 3,
                    InsertMode = Helpers.Collections.ItemInsertMode.Single,
                    Predicate = g => g is SeparatorTabItemsGroup separator && separator.Key == "Pinned-Default",
                    Comparer = defaultTabItemsGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Middle,
                    Predicate = g => g is TabItemsDefaultGroup,
                    Comparer = defaultTabItemsGroupComparer
                }
            };
            this.SortedTabItemsGroups = new Helpers.Collections.SortedObservableCollection<TabItemsGroupBase>(
                defaultTabItemsGroupComparer,
                priorityGroups
                );

            _tabItemsSelectionCoordinator = new Helpers.Collections.GroupsSelectionCoordinator<TabItemsGroupBase, TabItemBase>(this.SortedTabItemsGroups);
            _tabItemsSelectionCoordinator.OnItemSelectionChanged = this.OnTabItemSelectionChanged;
            _tabItemsSelectionCoordinator.OnSelectionStateChanged = this.OnSelectionStateChanged;

            _tabNavigationController = new Navigation.TabNavigationController(
                _tabItemsSelectionCoordinator,
                this.GetVisibleTabItems
            ) {
                SelectionActivationPolicy = TabSelectionActivationPolicy
            };
            _keyboardTabNavigationExtension = new Navigation.KeyboardTabNavigationExtension(_tabNavigationController);
            _keyboardTabNavigationExtension.FocusedItemChanged += this.OnKeyboardFocusedTabItemChanged;
            _keyboardTabNavigationExtension.InputTargetRestoreRequested += this.OnKeyboardInputTargetRestoreRequested;
            _tabNavigationController.AddExtension(_keyboardTabNavigationExtension);

            _textEditorOverlayController = new VsShell.TextEditor.Overlay.TextEditorOverlayController(PackageServices.Dte2);
        }

        private void UninitializeTabItemsSelectionCoordinator() {
            _keyboardTabNavigationExtension.FocusedItemChanged -= this.OnKeyboardFocusedTabItemChanged;
            _keyboardTabNavigationExtension.InputTargetRestoreRequested -= this.OnKeyboardInputTargetRestoreRequested;
            _textEditorOverlayController.Dispose();
        }


        // 
        // ░ BackgroundRoutine
        //
        private void InitializeBackgroundRoutine() {
            _tabsManagerStateTimer = new DispatcherTimer();
            _tabsManagerStateTimer.Interval = TimeSpan.FromSeconds(2);
            _tabsManagerStateTimer.Tick += this.TabsManagerStateTimerHandler;
            _tabsManagerStateTimer.Start();
        }

        private void UninitializeBackgroundRoutine() {
            if (_tabsManagerStateTimer == null) {
                return;
            }

            _tabsManagerStateTimer.Stop();
            _tabsManagerStateTimer.Tick -= this.TabsManagerStateTimerHandler;
            _tabsManagerStateTimer = null;
        }


        //
        // ░ Event handlers
        // ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 
        // 
        // ░ DTE
        //
        private void OnDocumentOpened(EnvDTE.Document document) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnDocumentOpened()");
            ThreadHelper.ThrowIfNotOnUIThread();
            
            // Log params:
            Helpers.Diagnostic.Logger.LogParam($"document.Name = {document?.Name}");

            var tabItemDocument = this.FindTabItem(document);
            if (tabItemDocument == null) {
                tabItemDocument = new TabItemDocument(document);
            }

            if (tabItemDocument.ShellDocument.IsDocumentInPreviewTab()) {
                this.AddDocumentToPreview(tabItemDocument);
            }
            else {
                this.AddTabItemToAutoDeterminedGroupIfMissing(tabItemDocument);
            }

        }


        private void OnDocumentSaved(EnvDTE.Document document) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnDocumentSaved()");
            ThreadHelper.ThrowIfNotOnUIThread();
            
            // Log params:
            Helpers.Diagnostic.Logger.LogParam($"document.Name = {document?.Name}");

            this.TabsManagerStateTimerHandler(null, null);
        }


        private void OnDocumentClosing(EnvDTE.Document document) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnDocumentClosing()");
            ThreadHelper.ThrowIfNotOnUIThread();

            // Log params:
            Helpers.Diagnostic.Logger.LogParam($"document.Name = {document?.Name}");

            if (!string.IsNullOrEmpty(document?.FullName)) {
                _textEditorOverlayController.OnDocumentClosing(document.FullName);
            }

            var tabItemDocument = this.FindTabItem(document);
            if (tabItemDocument != null) {
                // Если закрытие пришло не из CloseTabItems, оно было инициировано самой VS
                // (крестик, команда меню и т.п.) и должно попасть в историю отдельной операцией.
                if (!_isSolutionClosing && !_tabsBeingClosed.Remove(this.GetHistoryKey(tabItemDocument))) {
                    this.PushClosedTabsOperation(new[] { this.CreateClosedTabEntry(tabItemDocument) });
                }
                this.RemoveTabItemFromGroups(tabItemDocument);
            }
            else {
                Helpers.Diagnostic.Logger.LogWarning($"\"{document?.Name}\" not found in collections");
            }
        }
        

        private void OnWindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window lostFocus) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (gotFocus == null) {
                return;
            }

            var activatedShellWindow = new VsShell.Document.ShellWindow(gotFocus);
            if (!activatedShellWindow.IsTabWindow()) {
                return;
            }

            TabItemBase? tabItem;
            if (activatedShellWindow.Window.Document != null) {
                // Для уже известного документа не создаём временный TabItemDocument и его
                // служебное состояние при каждом переключении окна.
                tabItem = this.FindTabItem(activatedShellWindow.Window.Document.FullName);
                tabItem ??= this.AddTabItemToAutoDeterminedGroupIfMissing(new TabItemDocument(activatedShellWindow.Window.Document));
            }
            else {
                tabItem = this.FindTabItem(activatedShellWindow.Window);
                tabItem ??= this.AddTabItemToAutoDeterminedGroupIfMissing(new TabItemWindow(activatedShellWindow));
            }

            this.SelectActivatedTabItem(tabItem);

            if (tabItem is TabItemWindow) {
                this.UpdateWindowTabsInfo();
                this.SaveOpenToolWindows();
            }

            // IVsWindowFrameActivate не гарантирован при восстановлении сессии и повторной
            // активации того же frame. После DTE-события берём фактически активный editor view.
            if (tabItem is TabItemDocument) {
                VsixThreadHelper.RunOnUiThread(
                    Dispatcher,
                    this._textEditorOverlayController.Show,
                    DispatcherPriority.Background
                );
            }

        }


        private void OnWindowClosing(EnvDTE.Window closingWindow) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnWindowClosing()");
            ThreadHelper.ThrowIfNotOnUIThread();

            Helpers.Diagnostic.Logger.LogParam($"closingWindow.Caption = {closingWindow?.Caption}");
            var tabItemWindow = this.FindTabItem(closingWindow);
            // У tool window нет DocumentClosing, поэтому внешнее закрытие отслеживается здесь.
            // Для закрытия через панель ключ уже находится в _tabsBeingClosed.
            if (tabItemWindow != null && !_isSolutionClosing && !_tabsBeingClosed.Remove(this.GetHistoryKey(tabItemWindow))) {
                this.PushClosedTabsOperation(new[] { this.CreateClosedTabEntry(tabItemWindow) });
            }
            VsixThreadHelper.RunOnUiThread(Dispatcher, this.SaveOpenToolWindows, DispatcherPriority.Background);
        }


        private void OnSolutionClosing() {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnSolutionClosing()");
            ThreadHelper.ThrowIfNotOnUIThread();

            // Сохраняем снимок до того, как Visual Studio начнёт последовательно закрывать окна.
            // Иначе WindowClosing в конце завершения IDE перезапишет историю пустым списком.
            this.SaveOpenToolWindows();
            _isSolutionClosing = true;
            _closedTabsHistory.Clear();
            _tabsBeingClosed.Clear();

            // TODO: try replace with this.Unload()
            this.SortedTabItemsGroups.Clear();
            this.UninitializeFileWatcher();
        }
        

        // 
        // ░ FileWatcher
        //
        private void OnFileChanged(object sender, FileSystemEventArgs e) {
            //using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnFileChanged()");

            if (this.IsTemporaryFile(e.FullPath)) {
                return;
            }

            ThreadHelper.JoinableTaskFactory.Run(async () => {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                this.UpdateDocumentUI(e.FullPath);
            });
        }

        private void OnFileRenamed(object sender, RenamedEventArgs e) {
            //using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnFileRenamed()");

            if (this.IsTemporaryFile(e.FullPath) || this.IsTemporaryFile(e.OldFullPath)) {
                return;
            }

            ThreadHelper.JoinableTaskFactory.Run(async () => {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                this.UpdateDocumentUI(e.OldFullPath, e.FullPath);
            });
        }

        private void OnFileDeleted(object sender, FileSystemEventArgs e) {
            //using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnFileDeleted()");

            if (this.IsTemporaryFile(e.FullPath)) {
                return;
            }

            ThreadHelper.JoinableTaskFactory.Run(async () => {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                var documentFullName = e.FullPath;
                var tabItemDocument = this.FindTabItem(documentFullName);
                if (tabItemDocument != null) {
                    this.RemoveTabItemFromGroups(tabItemDocument);
                }
            });
        }


        //
        // ░ VsShellTrackers 
        //
        private void OnDocumentActivatedExternally(VsShell._EventArgs.DocumentNavigationEventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var tabItem = this.FindTabItem(e.CurrentDocumentFullName);
            if (tabItem != null) {
                // Навигация внутри редактора является внешней по отношению к панели:
                // синхронизируем и рамку реального frame, и selection вкладки.
                this.SetActiveFrameTabItem(tabItem);
                tabItem.Metadata.SetFlag("IsActivatedExternally", true);
                this.SelectActivatedTabItem(tabItem);
            }
        }


        private void OnVsWindowFrameActivated(IVsWindowFrame vsWindowFrame) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // NOTE: При запуске VS всегда активен последний документ сессии и
            //       OnVsWindowFrameActivated не сработает при переактивации этого же документа в ручную,
            //       поэтому явно ставим TextEditorFrameFocused = true в OnPreviewMouseDown.

            //VSFM_Dock = 0,     // ToolWindow: фиксировано сбоку, снизу и т.п. (Solution Explorer, Output и т.п.)
            //VSFM_MDIChild = 1, // Документное окно: занимает пространство текстового редактора, находится во вкладках
            //VSFM_Float = 2     // Плавающее окно: вытащено за пределы главного окна
            vsWindowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_FrameMode, out var mode);

            bool isMdiChild = mode != null && (VSFRAMEMODE)(int)mode == VSFRAMEMODE.VSFM_MdiChild;
            Helpers.GlobalFlags.SetFlag("TextEditorFrameFocused", isMdiChild);

            // IsFrameActive — отдельное состояние от IsSelected: только оно управляет фиолетовой
            // рамкой и должно изменяться вслед за реально активированным VS-фреймом, а не мультивыбором.
            var activeWindow = PackageServices.Dte2.ActiveWindow;
            TabItemBase activeFrameTabItem = null;
            if (activeWindow?.Document != null) {
                // Обычная документная вкладка сопоставляется по EnvDTE.Document.
                activeFrameTabItem = this.FindTabItem(activeWindow.Document);
            }
            else if (activeWindow != null) {
                // Встроенные страницы и tool windows документа не имеют и сопоставляются по Window.
                activeFrameTabItem = this.FindTabItem(activeWindow);
            }

            // Если окно не представлено в нашей панели, не назначаем рамку произвольной вкладке.
            if (activeFrameTabItem != null) {
                this.SetActiveFrameTabItem(activeFrameTabItem);
            }


            if (isMdiChild) {
                // Получаем содержимое активного окна (View)
                vsWindowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out var docView);

                // Если это кодовое окно, запрашиваем его "основной" текстовый редактор
                if (docView is IVsCodeWindow codeWindow) {
                    if (codeWindow.GetPrimaryView(out var textView) == VSConstants.S_OK && textView != null) {
                        // Это полноценный редактор
                        _textEditorOverlayController.ActivateEditorFrame(textView);
                        return;
                    }
                }
            }
            _textEditorOverlayController.Hide();
        }


        private void OnTextEditorNavigatedToDocument(VsShell._EventArgs.DocumentNavigationEventArgs e) {
            //if (e.PreviousDocumentFullName != null) {
            //    var ext = System.IO.Path.GetExtension(e.PreviousDocumentFullName);
            //    switch (ext) {
            //        case ".h":
            //        case ".hpp":
            //        //case ".cpp":
            //            break;

            //        default:
            //            return;
            //    }

            //    var fromTabItemDocument = this.FindTabItem(e.PreviousDocumentFullName);
            //    if (fromTabItemDocument == null) {
            //        return;
            //    }                

            //    var toTabItemDocument = this.FindTabItem(e.CurrentDocumentFullName);
            //    if (toTabItemDocument == null) {
            //        return;
            //    }


            //    this.MoveDocumentToProjectGroup(toTabItemDocument, fromTabItemDocument.SolutionProjectNodeContext);

            //    //var solutionHierarchyAnalyzer = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance;
            //    //var fromSourcesProjectNodes = solutionHierarchyAnalyzer.SourcesRepresentationsTable
            //    //    .GetProjectsByDocumentPath(fromTabItemDocument.FullName);

            //    //if (fromSourcesProjectNodes.Count > 0) {
            //    //    // NOTE: fromSourcesProjectNodes.Count usually == 1.
            //    //    //this.MoveDocumentToProjectGroup(toTabItemDocument, fromSourcesProjectNodes[0]);
            //    //    return;
            //    //}

            //    //var fromSharedItemsProjectNodes = solutionHierarchyAnalyzer.SharedItemsRepresentationsTable
            //    //    .GetProjectsByDocumentPath(fromTabItemDocument.FullName);

            //    //if (fromSharedItemsProjectNodes.Count > 0) {
            //    //    // TODO: add support .h (where fromProjectNodes.Count > 0).
            //    //    this.MoveDocumentToProjectGroup(toTabItemDocument, new TabItemProject(fromSharedItemsProjectNodes[0]));
            //    //    return;
            //    //}

            //    //var fromExternalIncludesProjectNodes = solutionHierarchyAnalyzer.ExternalIncludeRepresentationsTable
            //    //    .GetProjectsByDocumentPath(fromTabItemDocument.FullName);

            //    //if (fromExternalIncludesProjectNodes.Count > 0) {
            //    //    // TODO: add support .h (where fromProjectNodes.Count > 0).
            //    //    this.MoveDocumentToProjectGroup(toTabItemDocument, fromExternalIncludesProjectNodes));
            //    //    return;
            //    //}
            //}
        }


        // 
        // ░ TabItemsSelectionCoordinator
        //
        private void OnTabItemSelectionChanged(TabItemsGroupBase group, TabItemBase tabItem, bool isSelected) {
            // Выделение здесь намеренно не активирует документ. Решение об активации принимает
            // TabNavigationController с учётом модификаторов и выбранной TabSelectionActivationPolicy.
            // Флаг помечает одно внешнее изменение от VS; после доставки события его нужно погасить,
            // чтобы следующий пользовательский выбор обрабатывался как обычный.
            if (tabItem.Metadata.GetFlag("IsActivatedExternally")) {
                tabItem.Metadata.SetFlag("IsActivatedExternally", false);
            }
        }

        private void OnSelectionStateChanged(Helpers.Enums.SelectionState selectionState) {
            this.IsMultipleTabSelection = selectionState == Helpers.Enums.SelectionState.Multiple;
        }


        // 
        // ░ BackgroundRoutine
        //
        private void TabsManagerStateTimerHandler(object sender, EventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();
            // NOTE: Нужно использовать копии коллекций для безопасного перебора.
            // Поэтому в foreach вызывай .ToList() у коллекций.

            // === [A] Обновление статуса сохранения документов ===
            var openDocuments = PackageServices.Dte2.Documents
                .Cast<EnvDTE.Document>()
                .ToDictionary(document => document.FullName, StringComparer.OrdinalIgnoreCase);

            foreach (var tabItemsGroup in this.SortedTabItemsGroups.ToList()) {
                foreach (var tabItem in tabItemsGroup.Items.ToList()) {
                    if (openDocuments.TryGetValue(tabItem.FullName, out var document)) {
                        if (document.Saved) {
                            tabItem.Caption = tabItem.Caption.TrimEnd('*');
                        }
                        else {
                            if (!tabItem.Caption.EndsWith("*")) {
                                tabItem.Caption += "*";
                            }
                        }
                    }
                }
            }


            // === [B] Перемещение preview-документов в основную группу ===
            var previewGroup = this.SortedTabItemsGroups.FirstOrDefault(g => g is TabItemsPreviewGroup);
            if (previewGroup != null) {
                foreach (var tabItemDocument in previewGroup.Items.OfType<TabItemDocument>().ToList()) {
                    if (!tabItemDocument.ShellDocument.IsDocumentInPreviewTab()) {
                        this.MoveDocumentFromPreviewToMainGroup(tabItemDocument);
                    }
                }
            }


            // === [C] Удаление закрытых окон типа TabItemWindow ===
            var openWindowIds = new HashSet<string>();

            try {
                openWindowIds = PackageServices.Dte2.Windows
                    .Cast<EnvDTE.Window>() // Приводим COM-коллекцию к типизированной, чтобы использовать LINQ
                    .Select(w => VsShell.Document.ShellWindow.GetWindowId(w))
                    .Where(id => !string.IsNullOrEmpty(id))
                    .ToHashSet();
            }
            catch (Exception ex) {
                // Иногда обращение к _dte.Windows может выбросить COMException,особенно если окно закрывается
                // или COM-объект уже недоступен. Поэтому оборачиваем перебор в try-catch для устойчивости.
                Helpers.Diagnostic.Logger.LogError($"Failed to enumerate windows: {ex.Message}");
            }

            foreach (var group in this.SortedTabItemsGroups.ToList()) {
                var toRemove = group.Items
                    .OfType<TabItemWindow>()
                    // Новые встроенные страницы VS (например All Settings) могут быть
                    // доступны через событие WindowActivated, но отсутствовать в DTE.Windows.
                    // Пока такое окно видимо, отсутствие в COM-коллекции не означает закрытие.
                    .Where(w => !openWindowIds.Contains(w.WindowId) && !IsWindowVisible(w))
                    .ToList();

                foreach (var tab in toRemove) {
                    this.RemoveTabItemFromGroups(tab);
                }
            }

            // === [D] Обновление окон типа TabItemWindow ===
            this.UpdateWindowTabsInfo();
        }

        private static bool IsWindowVisible(TabItemWindow tabItemWindow) {
            ThreadHelper.ThrowIfNotOnUIThread();

            try {
                return tabItemWindow.ShellWindow.Window.Visible;
            }
            catch (Exception) {
                return false;
            }
        }


        //
        // ░ UI click handlers
        // ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 
        //
        // ░ Commands
        //
        private void OnPinTabItem(object parameter) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnPinTabItem()");
            ThreadHelper.ThrowIfNotOnUIThread();

            if (parameter is not TabItemBase tabItem) {
                return;
            }

            // Если вкладка уже закреплена — ничего не делаем
            if (tabItem.IsPinnedTab) {
                return;
            }

            // Ищем вкладку и группу, в которой она сейчас находится
            var current = this.FindTabItemWithGroup(tabItem);
            if (current == null) {
                return; // Не найдена — нечего обрабатывать
            }

            var item = current.Value.Item;
            var oldGroup = current.Value.Group;

            this.RemoveTabItemFromGroup(item, oldGroup);
            this.AddTabItemToGroupIfMissing(tabItem, new TabItemsPinnedGroup(oldGroup.GroupName));
        }


        private void OnUnpinTabItem(object parameter) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnUnpinTabItem()");
            ThreadHelper.ThrowIfNotOnUIThread();

            if (parameter is not TabItemBase tabItem) {
                return;
            }

            if (!tabItem.IsPinnedTab) {
                return;
            }

            var current = this.FindTabItemWithGroup(tabItem);
            if (current == null) {
                return;
            }

            var item = current.Value.Item;
            var oldGroup = current.Value.Group;

            this.RemoveTabItemFromGroup(item, oldGroup);
            this.AddTabItemToGroupIfMissing(tabItem, new TabItemsDefaultGroup(oldGroup.GroupName));
        }


        private void OnCloseTabItem(object parameter) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnCloseTabItem()");
            ThreadHelper.ThrowIfNotOnUIThread();

            if (parameter is not TabItemBase tabItem) {
                return;
            }

            var selectedItems = _tabItemsSelectionCoordinator.SelectedItems;
            bool closeSelection = selectedItems.Count > 1 && selectedItems.Any(entry => ReferenceEquals(entry.Item, tabItem));
            var itemsToClose = closeSelection
                ? selectedItems.Select(entry => entry.Item).ToList()
                : new List<TabItemBase> { tabItem };

            this.CloseTabItems(itemsToClose);
        }

        private void OnCloseSelectedTabItems(object parameter) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnCloseSelectedTabItems()");
            ThreadHelper.ThrowIfNotOnUIThread();

            var itemsToClose = _tabItemsSelectionCoordinator.SelectedItems
                .Select(entry => entry.Item)
                .ToList();

            this.CloseTabItems(itemsToClose);
        }

        private void CloseTabItems(IReadOnlyList<TabItemBase> tabItems) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Снимок создаётся до вызова DTE Close(): DocumentClosing приходит синхронно и сразу
            // удаляет TabItem вместе с информацией о его исходной группе.
            var entries = tabItems
                .Select(this.CreateClosedTabEntry)
                .ToList();
            if (entries.Count > 0) {
                this.PushClosedTabsOperation(entries);
                foreach (var tabItem in tabItems) {
                    // Помечаем всю пачку заранее: каждое последующее событие закрытия только
                    // погасит свой ключ и не раздробит одну операцию на несколько undo-шагов.
                    _tabsBeingClosed.Add(this.GetHistoryKey(tabItem));
                }
            }

            // Используем заранее созданный снимок: DocumentClosing синхронно удаляет элементы
            // из групп и одновременно перестраивает текущее выделение.
            foreach (var tabItem in tabItems) {
                try {
                    if (tabItem is TabItemDocument tabItemDocument) {
                        Helpers.Diagnostic.Logger.LogDebug($"close document \"{tabItemDocument.ShellDocument.Document.FullName}\"");
                        tabItemDocument.ShellDocument.Document.Close();
                        // Документ будет удалён через OnDocumentClosing.
                    }
                    else if (tabItem is TabItemWindow tabItemWindow) {
                        Helpers.Diagnostic.Logger.LogDebug($"close window \"{tabItemWindow.ShellWindow.Window.Caption}\"");
                        tabItemWindow.ShellWindow.Window.Close();

                        // Для tool window событие DocumentClosing не приходит.
                        this.RemoveTabItemFromGroups(tabItemWindow);
                    }
                }
                catch (Exception ex) {
                    _tabsBeingClosed.Remove(this.GetHistoryKey(tabItem));
                    Helpers.Diagnostic.Logger.LogError($"Failed to close tab '{tabItem.Caption}': {ex}");
                }
            }

            this.VirtualMenuControl.HideImmediately();
        }

        private void OnKeepOpenedTabItem(object parameter) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnKeepOpenedTabItem()");
            ThreadHelper.ThrowIfNotOnUIThread();

            if (parameter is TabItemBase tabItem) {
                if (tabItem is TabItemDocument tabItemDocument) {
                    this.MoveDocumentFromPreviewToMainGroup(tabItemDocument);
                    tabItemDocument.ShellDocument.OpenDocumentAsPinned();
                }
            }
        }

        private void OnCopyTabName(object parameter) {
            if (parameter is TabItemBase tabItem) {
                this.CopyTabTextToClipboard(tabItem.Caption, "name");
            }
        }

        private void OnCopyTabPath(object parameter) {
            if (parameter is TabItemBase tabItem) {
                this.CopyTabTextToClipboard(tabItem.FullName, "path");
            }
        }

        private void CopyTabTextToClipboard(string text, string valueKind) {
            ThreadHelper.ThrowIfNotOnUIThread();
            try {
                Clipboard.SetText(text ?? string.Empty);
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to copy tab {valueKind} to clipboard: {ex}");
            }
        }

        private void OnOpenLocationTabItem(object parameter) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnOpenTabLocation()");
            ThreadHelper.ThrowIfNotOnUIThread();

            if (parameter is TabItemBase tabItem) {
                if (tabItem is TabItemDocument tabItemDocument) {
                    try {
                        string filePath = tabItemDocument.FullName;

                        if (System.IO.File.Exists(filePath)) {
                            string args = $"/select,\"{filePath}\"";
                            System.Diagnostics.Process.Start("explorer.exe", args);
                        }
                        else {
                            Helpers.Diagnostic.Logger.LogWarning($"File not found: {filePath}");
                        }
                    }
                    catch (Exception ex) {
                        Helpers.Diagnostic.Logger.LogError($"Failed to open tab location: {ex.Message}");
                    }
                }

                this.VirtualMenuControl.HideImmediately();
            }
        }

        private void OnMoveTabItemToRelatedProject(object parameter) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnMoveTabItemToRelatedProject()");
            ThreadHelper.ThrowIfNotOnUIThread();
            this.VirtualMenuControl.HideImmediately();

            if (parameter is DocumentProjectReferencesInfo.RefEntry refEntry) {
                this.MoveDocumentToProjectGroup(refEntry.DocumentEntryBase);
            }
            else if (parameter is DocumentProjectReferencesInfo.GroupContextEntry groupContextEntry && groupContextEntry.CanSwitch) {
                // Меню может закрыться и пересобрать набор выделения уже после первого переноса,
                // поэтому обходим снимок ссылок, собранный при открытии меню.
                var activeDocumentPathBefore = PackageServices.Dte2.ActiveDocument?.FullName;
                var contextSwitchPlan = this.BuildGroupContextSwitchPlan(groupContextEntry.DocumentReferences);
                var activatedContextSwitchSources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var documentReference in groupContextEntry.DocumentReferences.ToList()) {
                    contextSwitchPlan.TryGetValue(documentReference, out var contextSwitchSourcePath);
                    this.MoveDocumentToProjectGroup(
                        documentReference.DocumentEntryBase,
                        playFeedback: false,
                        preferredContextSwitchSourcePath: contextSwitchSourcePath,
                        activatedContextSwitchSources: activatedContextSwitchSources
                    );
                }

                Console.Beep(frequency: 1000, duration: 300);

                // Каждый shared-header может запланировать своё переоткрытие. После завершения
                // всей пачки возвращаем фокус в документ, активный до выбора контекста.
                VsixThreadHelper.RunOnUiThread(async () => {
                    await Task.Delay(450);
                    var currentDocument = PackageServices.Dte2.Documents
                        .Cast<EnvDTE.Document>()
                        .FirstOrDefault(document => string.Equals(
                            document.FullName,
                            activeDocumentPathBefore,
                            StringComparison.OrdinalIgnoreCase
                        ));

                    if (currentDocument != null) {
                        currentDocument.Activate();
                    }
                    else if (!string.IsNullOrEmpty(activeDocumentPathBefore)) {
                        Helpers.Diagnostic.Logger.LogDebug($"[OnMoveTabItemToRelatedProject] Cannot restore active document '{activeDocumentPathBefore}' because its current frame was not found.");
                    }
                });
            }
        }

        private void OnMoveTabItemToRelatedProjectFile(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.VirtualMenuControl.HideImmediately();

            if (parameter is not ProjectContextSourceEntry sourceEntry) {
                return;
            }

            var projectEntry = sourceEntry.ProjectContext switch {
                DocumentProjectReferencesInfo.RefEntry refEntry => refEntry.ProjectEntry,
                DocumentProjectReferencesInfo.GroupContextEntry groupContext => groupContext.ProjectEntry,
                _ => null
            };
            if (projectEntry?.MultiState.Current is not VsShell.Project.LoadedProject loadedProject) {
                return;
            }

            var sourceDocument = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance
                .SourcesRepresentationsTable
                .GetDocumentByProjectAndDocumentPath(projectEntry.BaseViewModel, sourceEntry.SourcePath);
            if (sourceDocument?.BaseViewModel.HierarchyItemEntry.BaseViewModel is VsShell.Hierarchy.RealHierarchyItem hierarchyItem) {
                int hr = VsShell.Utils.VsHierarchyUtils.ClickOnSolutionHierarchyItem(
                    loadedProject.ProjectHierarchy.VsRealHierarchy,
                    hierarchyItem.ItemId
                );

                ErrorHandler.ThrowOnFailure(hr);
                return;
            }

            // Shared-файл может входить в граф конкретного project context, но не иметь
            // собственного узла в SourcesRepresentationsTable этого .vcxproj. В таком случае
            // открываем сам физический файл, не запуская переключение контекста хедера.
            if (!File.Exists(sourceEntry.SourcePath)) {
                Helpers.Diagnostic.Logger.LogWarning($"[OnMoveTabItemToRelatedProjectFile] File '{sourceEntry.SourcePath}' does not exist.");
                return;
            }

            Helpers.Diagnostic.Logger.LogDebug($"[OnMoveTabItemToRelatedProjectFile] Opening shared source by path: '{sourceEntry.SourcePath}'.");
            PackageServices.Dte2.ItemOperations
                .OpenFile(sourceEntry.SourcePath, EnvDTE.Constants.vsViewKindTextView)
                .Activate();
        }

        private void OnProjectContextIncludersOpen(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (parameter is not FrameworkElement anchor ||
                anchor.DataContext is not Helpers.MenuItemCommand menuItem ||
                menuItem.CommandParameterContext is not object projectContext) {

                return;
            }

            if (this.VirtualMenuControl.IsChildMenuOpen &&
                ReferenceEquals(this.VirtualMenuControl.CurrentChildMenuDataContext, projectContext)) {

                this.VirtualMenuControl.HideChild();
                return;
            }

            this.ShowProjectContextIncludersMenu(anchor, projectContext);
        }

        private void OnProjectContextMenuItemMouseEnter(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!this.VirtualMenuControl.IsChildMenuOpen ||
                parameter is not MenuItem menuItem ||
                menuItem.DataContext is not Helpers.MenuItemCommand menuItemCommand) {

                return;
            }

            object projectContext = menuItemCommand.CommandParameterContext;
            if (!ReferenceEquals(this.VirtualMenuControl.CurrentChildMenuDataContext, projectContext)) {
                this.ShowProjectContextIncludersMenu(menuItem, projectContext);
            }
        }

        private void ShowProjectContextIncludersMenu(FrameworkElement anchor, object projectContext) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var sourcePaths = this.GetProjectContextSwitchSourcePaths(projectContext);
            string projectName = projectContext switch {
                DocumentProjectReferencesInfo.RefEntry refEntry => refEntry.ProjectEntry.BaseViewModel.Name,
                DocumentProjectReferencesInfo.GroupContextEntry groupContext => groupContext.ProjectEntry.BaseViewModel.Name,
                _ => "Project"
            };

            var childItems = new ObservableCollection<Helpers.IMenuItem> {
                new Helpers.MenuItemHeader { Header = projectName }
            };

            if (sourcePaths.Count == 0) {
                childItems.Add(new Helpers.MenuItemHeader { Header = "No transitive including files" });
            }
            else {
                foreach (string sourcePath in sourcePaths) {
                    childItems.Add(new Helpers.MenuItemCommand {
                        Header = Path.GetFileName(sourcePath),
                        Command = new Helpers.RelayCommand<object>(this.OnMoveTabItemToRelatedProjectFile),
                        CommandParameterContext = new ProjectContextSourceEntry(projectContext, sourcePath)
                    });
                }
            }

            Point screenPoint = anchor.ex_ToDpiAwareScreen(new Point(anchor.ActualWidth + 8, 0));
            this.VirtualMenuControl.ShowChild(screenPoint, projectContext, childItems);
        }

        private IReadOnlyList<string> GetProjectContextSwitchSourcePaths(object projectContext) {
            ThreadHelper.ThrowIfNotOnUIThread();

            IEnumerable<DocumentProjectReferencesInfo.RefEntry> references = projectContext switch {
                DocumentProjectReferencesInfo.RefEntry refEntry => new[] { refEntry },
                DocumentProjectReferencesInfo.GroupContextEntry groupContext => groupContext.DocumentReferences,
                _ => Array.Empty<DocumentProjectReferencesInfo.RefEntry>()
            };

            return references
                .SelectMany(reference => GetProjectContextSwitchSourcePaths(reference))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
                .ThenBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static VsShell.Document.Document? GetCurrentProjectContextDocument(VsShell.Document.DocumentEntryBase entry) {
            return entry switch {
                VsShell.Document.ExternalIncludeEntry externalIncludeEntry =>
                    externalIncludeEntry.MultiState.Current as VsShell.Document.Document,
                VsShell.Document.SharedItemEntry sharedItemEntry =>
                    sharedItemEntry.MultiState.Current as VsShell.Document.Document,
                _ => null
            };
        }

        private static IReadOnlyList<string> GetProjectContextSwitchSourcePaths(DocumentProjectReferencesInfo.RefEntry reference) {
            var document = GetCurrentProjectContextDocument(reference.DocumentEntryBase);
            var targetProject = reference.ProjectEntry.MultiState.Current as VsShell.Project.LoadedProject;
            return document != null && targetProject != null
                ? document.GetProjectContextSwitchSourcePaths(targetProject)
                : Array.Empty<string>();
        }

        private IReadOnlyDictionary<DocumentProjectReferencesInfo.RefEntry, string> BuildGroupContextSwitchPlan(
            IReadOnlyList<DocumentProjectReferencesInfo.RefEntry> documentReferences
            ) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var candidatesByDocument = documentReferences.ToDictionary(
                reference => reference,
                reference => GetProjectContextSwitchSourcePaths(reference).ToHashSet(StringComparer.OrdinalIgnoreCase)
            );

            var openDocumentPaths = PackageServices.Dte2.Documents
                .Cast<EnvDTE.Document>()
                .Select(document => document.FullName)
                .Where(path => !string.IsNullOrEmpty(path))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var uncoveredDocuments = candidatesByDocument
                .Where(pair => pair.Value.Count > 0)
                .Select(pair => pair.Key)
                .ToHashSet();
            var result = new Dictionary<DocumentProjectReferencesInfo.RefEntry, string>();

            // Жадное покрытие: на каждом шаге берём .cpp, который включает максимум ещё не
            // обработанных хедеров. Например, если Engine.cpp включает A.h, B.h и C.h, а
            // Tools.cpp только C.h, для всех трёх будет выбран один Engine.cpp.
            while (uncoveredDocuments.Count > 0) {
                var selectedSourcePath = uncoveredDocuments
                    .SelectMany(reference => candidatesByDocument[reference])
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Select(path => new {
                        Path = path,
                        CoveredCount = uncoveredDocuments.Count(reference => candidatesByDocument[reference].Contains(path)),
                        IsOpen = openDocumentPaths.Contains(path)
                    })
                    .OrderByDescending(candidate => candidate.CoveredCount)
                    .ThenBy(candidate => candidate.IsOpen)
                    .ThenBy(candidate => candidate.Path, StringComparer.OrdinalIgnoreCase)
                    .First()
                    .Path;

                var coveredDocuments = uncoveredDocuments
                    .Where(reference => candidatesByDocument[reference].Contains(selectedSourcePath))
                    .ToList();
                foreach (var coveredDocument in coveredDocuments) {
                    result[coveredDocument] = selectedSourcePath;
                    uncoveredDocuments.Remove(coveredDocument);
                }
            }

            foreach (var referenceWithoutSource in candidatesByDocument.Where(pair => pair.Value.Count == 0)) {
                Helpers.Diagnostic.Logger.LogWarning($"[BuildGroupContextSwitchPlan] No transitive .cpp was found for '{referenceWithoutSource.Key.DocumentEntryBase.BaseViewModel.HierarchyItemEntry.BaseViewModel.FilePath}'. Project context will be changed without translation-unit activation.");
            }

            return result;
        }


        private void OnReloadDocumentReferencesProjects(object parameter) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnReloadDocumentReferencesProjects()");
            ThreadHelper.ThrowIfNotOnUIThread();

            if (parameter is DocumentProjectReferencesInfo documentProjectReferencesInfo) {
                foreach (var refEntry in documentProjectReferencesInfo.References) {
                    var projectViewModel = refEntry.ProjectEntry.BaseViewModel;
                    if (projectViewModel is VsShell.Project.UnloadedProject) {
                        VsShell.Utils.VsHierarchyUtils.ReloadProject(projectViewModel.ProjectGuid);
                    }
                }
            }
        }


        // 
        // ░ ContextMenu
        //
        private void OnTabItemContextMenuOpen(object parameter) {
            if (parameter is Controls.MenuControl.MenuOpeningArgs contextMenuOpeningArgs) {
                if (contextMenuOpeningArgs.DataContext is TabItemBase tabItem) {

                    switch (_tabItemsSelectionCoordinator.SelectionState) {
                        case Helpers.Enums.SelectionState.Single:
                            if (tabItem is TabItemDocument tabItemDocument) {
                                tabItemDocument.Metadata.SetFlag("IsCtxMenuOpenned", true);

                                this.ContextMenuItems = new ObservableCollection<Helpers.IMenuItem> {
                                    new Helpers.MenuItemCommand {
                                        Header = State.Constants.UI.CopyTabName,
                                        Command = new Helpers.RelayCommand<object>(this.OnCopyTabName),
                                        CommandParameterContext = contextMenuOpeningArgs.DataContext,
                                    },
                                    new Helpers.MenuItemCommand {
                                        Header = State.Constants.UI.CopyTabPath,
                                        Command = new Helpers.RelayCommand<object>(this.OnCopyTabPath),
                                        CommandParameterContext = contextMenuOpeningArgs.DataContext,
                                    },
                                    new Helpers.MenuItemSeparator(),
                                    new Helpers.MenuItemCommand {
                                        Header = State.Constants.UI.OpenTabLocation,
                                        Command = new Helpers.RelayCommand<object>(this.OnOpenLocationTabItem),
                                        CommandParameterContext = contextMenuOpeningArgs.DataContext,
                                    },
                                    new Helpers.MenuItemSeparator(),
                                    new Helpers.MenuItemCommand {
                                        Header = State.Constants.UI.CloseTab,
                                        Command = new Helpers.RelayCommand<object>(this.OnCloseTabItem),
                                        CommandParameterContext = contextMenuOpeningArgs.DataContext,
                                    },
                                };
                            }
                            else if (tabItem is TabItemWindow tabItemWindow) {
                                this.ContextMenuItems = new ObservableCollection<Helpers.IMenuItem> {
                                    new Helpers.MenuItemCommand {
                                        Header = State.Constants.UI.CopyTabName,
                                        Command = new Helpers.RelayCommand<object>(this.OnCopyTabName),
                                        CommandParameterContext = contextMenuOpeningArgs.DataContext,
                                    },
                                    new Helpers.MenuItemCommand {
                                        Header = State.Constants.UI.CopyTabPath,
                                        Command = new Helpers.RelayCommand<object>(this.OnCopyTabPath),
                                        CommandParameterContext = contextMenuOpeningArgs.DataContext,
                                    },
                                    new Helpers.MenuItemSeparator(),
                                    new Helpers.MenuItemCommand {
                                        Header = State.Constants.UI.CloseTab,
                                        Command = new Helpers.RelayCommand<object>(this.OnCloseTabItem),
                                        CommandParameterContext = contextMenuOpeningArgs.DataContext,
                                    },
                                };
                            }
                            break;

                        case Helpers.Enums.SelectionState.Multiple:
                            bool isTabItemAmongSelectedItems = _tabItemsSelectionCoordinator.SelectedItems
                                .Any(entry => ReferenceEquals(entry.Item, tabItem));

                            if (isTabItemAmongSelectedItems) {
                                this.ContextMenuItems = new ObservableCollection<Helpers.IMenuItem> {
                                    new Helpers.MenuItemCommand {
                                        Header = State.Constants.UI.CloseSelectedTabs,
                                        Command = new Helpers.RelayCommand<object>(this.OnCloseSelectedTabItems),
                                        CommandParameterContext = contextMenuOpeningArgs.DataContext,
                                    }
                                };
                            }
                            else {
                                contextMenuOpeningArgs.ShouldOpen = false;
                                tabItem.IsSelected = true;
                            }
                            break;
                    }
                }
            }
        }

        private void OnTabItemContextMenuClosed(object parameter) {
            if (parameter is Controls.MenuControl.MenuClosedArgs contextMenuClosedArgs) {
                if (contextMenuClosedArgs.DataContext is TabItemBase tabItem) {
                    tabItem.Metadata.SetFlag("IsCtxMenuOpenned", false);
                }
            }
        }


        // 
        // ░ VirtualMenu
        //
        private void CloseButton_MouseEnter(object sender, MouseEventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (sender is Button closeButton) {
                var tabItemControl = Helpers.VisualTree.FindParentByType<TabItemControl>(closeButton);
                if (tabItemControl == null) {
                    return;
                }

                // Находим родительский ListViewItem (где привязаны данные)
                var listViewItem = Helpers.VisualTree.FindParentByType<ListViewItem>(tabItemControl);
                if (listViewItem == null) {
                    return;
                }

                // Получаем привязанный объект (TabItemDocument)
                if (listViewItem.DataContext is TabItemDocument tabItemDocument) {
                    if (this.VirtualMenuControl.CurrentMenuDataContext is TabItemDocument previousTabItemDocument) {
                        previousTabItemDocument.Metadata.SetFlag("IsVirtualMenuOpenned", false);
                    }
                    // Таймер запускается только над крестиком. После открытия меню остаётся
                    // видимым при перемещении по вкладке и закрывается на MouseLeave всей строки.
                    var screenPoint = tabItemControl.ex_ToDpiAwareScreen(new Point(tabItemControl.ActualWidth + 20, -60));
                    this.VirtualMenuControl.Show(screenPoint, tabItemDocument);
                }
            }
        }

        private void InteractiveArea_MouseEnter(object sender, MouseEventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Вся площадь вкладки переключает контекст только у уже открытого меню.
            // Первичный показ по-прежнему запускается исключительно наведением на крестик.
            if (!this.VirtualMenuControl.IsMenuOpen ||
                sender is not TabItemControl tabItemControl ||
                tabItemControl.DataContext is not TabItemDocument tabItemDocument ||
                ReferenceEquals(this.VirtualMenuControl.CurrentMenuDataContext, tabItemDocument)) {
                return;
            }

            if (this.VirtualMenuControl.CurrentMenuDataContext is TabItemDocument previousTabItemDocument) {
                previousTabItemDocument.Metadata.SetFlag("IsVirtualMenuOpenned", false);
            }

            // Show обновляет уже открытый popup без задержки и отменяет таймер скрытия,
            // запущенный MouseLeave предыдущей вкладки.
            var screenPoint = tabItemControl.ex_ToDpiAwareScreen(new Point(tabItemControl.ActualWidth + 20, -60));
            this.VirtualMenuControl.Show(screenPoint, tabItemDocument);
        }

        private void InteractiveArea_MouseLeave(object sender, MouseEventArgs e) {
            //using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope($"InteractiveArea_MouseLeave()");
            ThreadHelper.ThrowIfNotOnUIThread();

            this.VirtualMenuControl.Hide();
        }

        private void OnTabItemVirtualMenuOpen(object parameter) {
            //using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("OnTabItemVirtualMenuOpen()");
            ThreadHelper.ThrowIfNotOnUIThread();
            this.VirtualMenuControl.HideChild();

            if (parameter is Controls.MenuControl.MenuOpeningArgs virtualMenuOpeningArgs) {
                if (virtualMenuOpeningArgs.DataContext is TabItemBase tabItem) {
                    
                    if (tabItem is TabItemDocument tabItemDocument) {
                        tabItemDocument.Metadata.SetFlag("IsVirtualMenuOpenned", true);

                        var newMenuItems = new List<Helpers.IMenuItem> {
                            new Helpers.MenuItemHeader {
                                Header = tabItem.Caption,
                            },
                            new Helpers.MenuItemCommand {
                                Header = State.Constants.UI.CopyTabName,
                                Command = new Helpers.RelayCommand<object>(this.OnCopyTabName),
                                CommandParameterContext = virtualMenuOpeningArgs.DataContext,
                            },
                            new Helpers.MenuItemCommand {
                                Header = State.Constants.UI.CopyTabPath,
                                Command = new Helpers.RelayCommand<object>(this.OnCopyTabPath),
                                CommandParameterContext = virtualMenuOpeningArgs.DataContext,
                            },
                            new Helpers.MenuItemSeparator(),
                            new Helpers.MenuItemCommand {
                                Header = State.Constants.UI.OpenTabLocation,
                                Command = new Helpers.RelayCommand<object>(this.OnOpenLocationTabItem),
                                CommandParameterContext = virtualMenuOpeningArgs.DataContext,
                            },
                            new Helpers.MenuItemCommand {
                                Header = State.Constants.UI.CloseTab,
                                Command = new Helpers.RelayCommand<object>(this.OnCloseTabItem),
                                CommandParameterContext = virtualMenuOpeningArgs.DataContext,
                            },
                        };
                        if (tabItemDocument.ShellDocument != null) {
                            if (this.TryGetSelectedHeaderGroup(tabItemDocument, out var selectedHeaders)) {
                                var groupContexts = this.BuildGroupContextEntries(selectedHeaders);

                                if (groupContexts.Count > 0) {
                                    newMenuItems.Add(new Helpers.MenuItemSeparator());

                                    var commonContexts = groupContexts.Where(context => context.IsAvailableForAll).ToList();
                                    var differingContexts = groupContexts.Where(context => !context.IsAvailableForAll).ToList();

                                    foreach (var groupContext in commonContexts) {
                                        newMenuItems.Add(new Helpers.MenuItemCommand {
                                            Header = groupContext.ProjectEntry.BaseViewModel.Name,
                                            Command = new Helpers.RelayCommand<object>(this.OnMoveTabItemToRelatedProject),
                                            CommandParameterContext = groupContext,
                                        });
                                    }

                                    if (commonContexts.Count > 0 && differingContexts.Count > 0) {
                                        newMenuItems.Add(new Helpers.MenuItemSeparator());
                                    }

                                    foreach (var groupContext in differingContexts) {
                                        newMenuItems.Add(new Helpers.MenuItemCommand {
                                            Header = groupContext.ProjectEntry.BaseViewModel.Name,
                                            Command = new Helpers.RelayCommand<object>(this.OnMoveTabItemToRelatedProject),
                                            CommandParameterContext = groupContext,
                                        });
                                    }
                                }
                            }
                            else {
                                var documentReferences = tabItemDocument.DocumentProjectReferencesInfo.GetAvailableReferences();

                                if (documentReferences.Count > 0) {
                                    newMenuItems.Add(new Helpers.MenuItemSeparator());

                                    foreach (var refEntry in documentReferences) {
                                        newMenuItems.Add(new Helpers.MenuItemCommand {
                                            Header = refEntry.ProjectEntry.BaseViewModel.Name,
                                            Command = new Helpers.RelayCommand<object>(this.OnMoveTabItemToRelatedProject),
                                            CommandParameterContext = refEntry,
                                        });
                                    }

                                    newMenuItems.Add(new Helpers.MenuItemCommand {
                                        Header = "Reload projects",
                                        Command = new Helpers.RelayCommand<object>(this.OnReloadDocumentReferencesProjects),
                                        CommandParameterContext = tabItemDocument.DocumentProjectReferencesInfo,
                                    });
                                }
                            }

                            this.UpdateVirtualMenuItems(newMenuItems);
                        }
                    }
                    else if (tabItem is TabItemWindow tabItemWindow) {
                        virtualMenuOpeningArgs.ShouldOpen = false;
                    }
                }
            }
        }

        private void UpdateVirtualMenuItems(IReadOnlyList<Helpers.IMenuItem> newItems) {
            // Сохраняем существующие модели и визуальные контейнеры там, где структура меню
            // совпадает. Обычно при переходе между вкладками меняются только Header и параметры.
            int commonCount = Math.Min(this.VirtualMenuItems.Count, newItems.Count);
            for (int index = 0; index < commonCount; index++) {
                var currentItem = this.VirtualMenuItems[index];
                var newItem = newItems[index];

                if (currentItem is Helpers.MenuItemHeader currentHeader &&
                    newItem is Helpers.MenuItemHeader newHeader) {
                    currentHeader.Header = newHeader.Header;
                    continue;
                }

                if (currentItem is Helpers.MenuItemSeparator && newItem is Helpers.MenuItemSeparator) {
                    continue;
                }

                if (currentItem is Helpers.MenuItemCommand currentCommand &&
                    newItem is Helpers.MenuItemCommand newCommand &&
                    GetVirtualMenuItemKind(currentCommand) == GetVirtualMenuItemKind(newCommand)) {
                    currentCommand.Header = newCommand.Header;
                    currentCommand.CommandParameterContext = newCommand.CommandParameterContext;
                    continue;
                }

                this.VirtualMenuItems[index] = newItem;
            }

            while (this.VirtualMenuItems.Count > newItems.Count) {
                this.VirtualMenuItems.RemoveAt(this.VirtualMenuItems.Count - 1);
            }

            for (int index = commonCount; index < newItems.Count; index++) {
                this.VirtualMenuItems.Add(newItems[index]);
            }
        }

        private static Type GetVirtualMenuItemKind(Helpers.MenuItemCommand menuItem) {
            return menuItem.CommandParameterContext switch {
                DocumentProjectReferencesInfo.RefEntry => typeof(DocumentProjectReferencesInfo.RefEntry),
                DocumentProjectReferencesInfo.GroupContextEntry => typeof(DocumentProjectReferencesInfo.GroupContextEntry),
                DocumentProjectReferencesInfo => typeof(DocumentProjectReferencesInfo),
                _ => typeof(TabItemBase)
            };
        }

        private bool TryGetSelectedHeaderGroup(TabItemDocument anchorTabItem, out IReadOnlyList<TabItemDocument> selectedHeaders) {
            var selectedItems = _tabItemsSelectionCoordinator.SelectedItems
                .Select(entry => entry.Item)
                .ToList();

            bool anchorIsSelected = selectedItems.Any(item => ReferenceEquals(item, anchorTabItem));
            bool allSelectedItemsAreHeaders = selectedItems.All(item =>
                item is TabItemDocument document && IsHeaderPath(document.FullName));

            if (_tabItemsSelectionCoordinator.SelectionState != Helpers.Enums.SelectionState.Multiple ||
                !anchorIsSelected ||
                selectedItems.Count < 2 ||
                !allSelectedItemsAreHeaders) {
                selectedHeaders = Array.Empty<TabItemDocument>();
                return false;
            }

            selectedHeaders = selectedItems.Cast<TabItemDocument>().ToList();
            return true;
        }

        private IReadOnlyList<DocumentProjectReferencesInfo.GroupContextEntry> BuildGroupContextEntries(
            IReadOnlyList<TabItemDocument> selectedHeaders
            ) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var referencesByDocument = selectedHeaders
                .Select(header => header.DocumentProjectReferencesInfo
                    .GetAvailableReferences(includeSingleProject: true)
                    .GroupBy(reference => reference.ProjectEntry.BaseViewModel.ProjectGuid)
                    .ToDictionary(group => group.Key, group => group.First()))
                .ToList();

            var allProjectGuids = referencesByDocument
                .SelectMany(references => references.Keys)
                .Distinct()
                .ToList();

            var result = new List<DocumentProjectReferencesInfo.GroupContextEntry>();
            foreach (var projectGuid in allProjectGuids) {
                var documentReferences = new List<DocumentProjectReferencesInfo.RefEntry>();
                foreach (var references in referencesByDocument) {
                    if (references.TryGetValue(projectGuid, out var documentReference)) {
                        documentReferences.Add(documentReference);
                    }
                }

                bool isAvailableForAll = documentReferences.Count == selectedHeaders.Count;
                result.Add(new DocumentProjectReferencesInfo.GroupContextEntry(
                    documentReferences[0].ProjectEntry,
                    documentReferences,
                    isAvailableForAll
                ));
            }

            return result
                .OrderByDescending(context => context.IsAvailableForAll)
                .ThenBy(context => context.ProjectEntry.BaseViewModel.UniqueName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool IsHeaderPath(string path) {
            string extension = Path.GetExtension(path);
            return
                string.Equals(extension, ".h", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".hpp", StringComparison.OrdinalIgnoreCase);
        }

        private void OnTabItemVirtualMenuClosed(object parameter) {
            this.VirtualMenuControl.HideChild();

            if (parameter is Controls.MenuControl.MenuClosedArgs virtualMenuClosedArgs) {
                if (virtualMenuClosedArgs.DataContext is TabItemBase tabItem) {
                    tabItem.Metadata.SetFlag("IsVirtualMenuOpenned", false);
                }
            }
        }

        private sealed class ProjectContextSourceEntry {
            public object ProjectContext { get; }
            public string SourcePath { get; }

            public ProjectContextSourceEntry(object projectContext, string sourcePath) {
                this.ProjectContext = projectContext;
                this.SourcePath = sourcePath;
            }
        }


        // 
        // ░ ScaleSelectorControl(s)
        //
        private void ApplyScaleUI() {
            if (this.DocumentScaleTransform != null) {
                this.DocumentScaleTransform.ScaleX = this.ScaleFactorUI;
                this.DocumentScaleTransform.ScaleY = this.ScaleFactorUI;
            }
        }

        private void ApplyScaleTabsCompactness() {
            // Новая база 100% равна прежней высоте при 120%: 22 * 1.2 = 26.4.
            Helpers.BaseUserControlResourceHelper.UpdateDynamicResource(this, "AppTabItemHeight", this.ScaleFactorTabsCompactness * 26.4);

            // Вкладки материализуются позже родительского контрола и получают отдельные экземпляры
            // ResourceDictionary, поэтому обновляем уже созданные элементы также напрямую.
            foreach (var tabItemControl in _tabItemControls) {
                this.ApplyScaleToTabItemControl(tabItemControl);
            }
        }


        //
        // ░ Internal logic
        // ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 
        //
        // ░ Adding tabs
        //
        private void LoadOpenDocuments() {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("LoadOpenDocuments()");
            ThreadHelper.ThrowIfNotOnUIThread();

            this.SortedTabItemsGroups.ToList();
            this.SortedTabItemsGroups.Clear();
            this.RestoreToolWindows();

            foreach (EnvDTE.Document document in PackageServices.Dte2.Documents) {
                var tabItemDocument = new TabItemDocument(document);
                this.AddTabItemToAutoDeterminedGroupIfMissing(tabItemDocument);
            }

            foreach (EnvDTE.Window window in PackageServices.Dte2.Windows) {
                if (window.Document != null) {
                    continue; // skip documents
                }

                var shellWindow = new VsShell.Document.ShellWindow(window);
                if (!shellWindow.IsTabWindow()) {
                    continue; // skip non tab windows
                }

                var tabItemWindow = new TabItemWindow(shellWindow);
                this.AddTabItemToAutoDeterminedGroupIfMissing(tabItemWindow);
            }

            this.SyncActiveDocumentWithPrimaryTabItem();

            VsixThreadHelper.RunOnUiThread(Dispatcher, () => {
                if (VsShell.TextEditor.TextEditorControlHelper.IsEditorActive()) {
                    Helpers.GlobalFlags.SetFlag("TextEditorFrameFocused", true);
                    _textEditorOverlayController.Show();
                }
            }, DispatcherPriority.Background);
        }

        private void RestoreToolWindows() {
            ThreadHelper.ThrowIfNotOnUIThread();

            var uiShell = Package.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
            if (uiShell == null) {
                return;
            }

            _isRestoringToolWindows = true;
            try {
                foreach (var windowId in Configuration.TabsManagerConfigurationService.OpenToolWindowIds) {
                    if (!Guid.TryParse(windowId, out var persistenceGuid)) {
                        continue;
                    }

                    try {
                        var result = uiShell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fForceCreate, ref persistenceGuid, out var frame);
                        if (ErrorHandler.Succeeded(result)) {
                            frame?.Show();
                        }
                    }
                    catch (Exception ex) {
                        Helpers.Diagnostic.Logger.LogWarning($"Failed to restore tool window '{windowId}': {ex.Message}");
                    }
                }
            }
            finally {
                _isRestoringToolWindows = false;
            }
        }

        private void SaveOpenToolWindows() {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_isRestoringToolWindows || _isSolutionClosing) {
                return;
            }

            var windowIds = PackageServices.Dte2.Windows
                .Cast<EnvDTE.Window>()
                // DTE.Windows содержит также закрытые скрытые окна. Без проверки Visible
                // они снова сохраняются и принудительно открываются при следующем запуске.
                .Where(window => window.Visible && window.Document == null && VsShell.Document.ShellWindow.IsTabWindow(window))
                .Select(VsShell.Document.ShellWindow.GetWindowId);

            Configuration.TabsManagerConfigurationService.SetOpenToolWindowIds(windowIds);
        }


        private void AddDocumentToPreview(TabItemDocument tabItemDocument) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("AddDocumentToPreview()");
            ThreadHelper.ThrowIfNotOnUIThread();

            // No need remove old because preview tab item group
            // guarded with ItemInsertMode == SingleWithReplaceExisting.

            var addedOrExistTabItem = this.AddTabItemToGroupIfMissing(tabItemDocument, new TabItemsPreviewGroup());
            addedOrExistTabItem.IsSelected = true;
        }


        private TabItemBase AddTabItemToAutoDeterminedGroupIfMissing(TabItemBase tabItem) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("AddTabItemToAutoDeterminedGroupIfMissing()");
            ThreadHelper.ThrowIfNotOnUIThread();

            TabItemsGroupBase tabItemGroup = null;

            if (tabItem is TabItemDocument tabItemDocument) {
                tabItemGroup = new TabItemsDefaultGroup(tabItemDocument.ShellDocument.GetDocumentProjectName());
            }
            else if (tabItem is TabItemWindow) {
                tabItemGroup = new TabItemsDefaultGroup("[Tool Windows]");
            }
            else {
                tabItemGroup = new TabItemsDefaultGroup("Other");
            }

            return this.AddTabItemToGroupIfMissing(tabItem, tabItemGroup);
        }


        private TabItemBase AddTabItemToGroupIfMissing(TabItemBase tabItem, TabItemsGroupBase tabItemGroup) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("AddTabItemToGroupIfMissing()");
            ThreadHelper.ThrowIfNotOnUIThread();

            Helpers.Diagnostic.Logger.LogParam($"tabItem.Caption = {tabItem?.Caption}");
            Helpers.Diagnostic.Logger.LogParam($"tabItemGroup.GroupName = {tabItemGroup?.GroupName}");

            var existingTabItem = this.FindTabItem(tabItem);
            if (existingTabItem != null) {
                return existingTabItem;
            }

            var existingGroup = this.SortedTabItemsGroups
                .FirstOrDefault(g => g.GetType() == tabItemGroup.GetType() && g.GroupName == tabItemGroup.GroupName);

            if (existingGroup == null) {
                this.SortedTabItemsGroups.Add(tabItemGroup);
                this.UpdateSeparatorsBetweenGroups();
                existingGroup = tabItemGroup;
            }

            // Evaluate properties:
            tabItem.IsPinnedTab = tabItemGroup is TabItemsPinnedGroup;

            if (tabItem is TabItemDocument tabItemDocument) {
                tabItemDocument.IsPreviewTab = tabItemGroup is TabItemsPreviewGroup;

                //var solutionHierarchyAnalyzer = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance;
                //var targetProjectNode = solutionHierarchyAnalyzer.ProjectNodes
                //    .FirstOrDefault(p => String.Equals(p.Name, tabItemGroup.GroupName, StringComparison.OrdinalIgnoreCase));

                //tabItemDocument.ProjectNodeContext = targetProjectNode;
            }

            Helpers.Diagnostic.Logger.LogDebug($"Added tab \"{tabItem.Caption}\" to group \"{tabItemGroup.GroupName}\": {tabItem}");
            existingGroup.Items.Add(tabItem);
            return tabItem;
        }


        // 
        // ░ Removing tabs
        //
        private void MoveDocumentFromPreviewToMainGroup(TabItemDocument tabItemDocument) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("MoveDocumentFromPreviewToMainGroup()");

            // Log params:
            Helpers.Diagnostic.Logger.LogParam($"tabItemDocument.ShellDocument.Document.Name = {tabItemDocument?.ShellDocument.Document.Name}");

            if (tabItemDocument == null) {
                return;
            }
            if (!tabItemDocument.IsPreviewTab) {
                return;
            }

            var previewGroup = this.SortedTabItemsGroups.FirstOrDefault(g => g is TabItemsPreviewGroup);
            if (previewGroup != null) {
                this.RemoveTabItemsGroup(previewGroup);
            }
            this.AddTabItemToAutoDeterminedGroupIfMissing(tabItemDocument);
        }


        private void MoveDocumentToProjectGroup(
            VsShell.Document.DocumentEntryBase documentEntryBase,
            bool playFeedback = true,
            string? preferredContextSwitchSourcePath = null,
            ISet<string>? activatedContextSwitchSources = null
            ) {
            using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("MoveDocumentToProjectGroup()");
            ThreadHelper.ThrowIfNotOnUIThread();
            
            this.OpenTabItemWithProjectContext(
                documentEntryBase,
                playFeedback,
                preferredContextSwitchSourcePath,
                activatedContextSwitchSources
            );

            var docVM = documentEntryBase.BaseViewModel;
            var hierarchyVM = documentEntryBase.BaseViewModel.HierarchyItemEntry.BaseViewModel;

            var tabItemDocument = this.FindTabItem(hierarchyVM.FilePath);
            if (tabItemDocument != null) {
                this.RemoveTabItemFromGroups(tabItemDocument);
                this.AddTabItemToGroupIfMissing(tabItemDocument, new TabItemsDefaultGroup(docVM.ProjectBaseViewModel.Name));
            }
        }


        private void OpenTabItemWithProjectContext(
            VsShell.Document.DocumentEntryBase documentEntryBase,
            bool playFeedback,
            string? preferredContextSwitchSourcePath = null,
            ISet<string>? activatedContextSwitchSources = null
            ) {
            if (documentEntryBase is VsShell.Document.ExternalIncludeEntry externalIncludeEntry) {
                if (externalIncludeEntry.MultiState.Current is VsShell.Document.ExternalInclude externalInclude) {
                    externalInclude.OpenWithProjectContext(
                        preferredContextSwitchSourcePath,
                        activatedContextSwitchSources
                    );
                    if (playFeedback) {
                        Console.Beep(frequency: 1000, duration: 300);
                    }
                }
                else if (externalIncludeEntry.MultiState.Current is VsShell.Document.InvalidatedDocument invalidatedDocument) {
                    System.Diagnostics.Debugger.Break();
                }
                return;
            }
            else if (documentEntryBase is VsShell.Document.SharedItemEntry sharedItemEntry) {
                if (sharedItemEntry.MultiState.Current is VsShell.Document.SharedItem sharedItem) {
                    sharedItem.OpenWithProjectContext(
                        preferredContextSwitchSourcePath,
                        activatedContextSwitchSources
                    );
                    if (playFeedback) {
                        Console.Beep(frequency: 1000, duration: 300);
                    }
                }
                else if (sharedItemEntry.MultiState.Current is VsShell.Document.InvalidatedDocument invalidatedDocument) {
                    invalidatedDocument.OpenWithProjectContext();
                }
                return;
            }
        }

        private void RemoveTabItemFromGroups(TabItemBase tabItem) {
            foreach (var group in this.SortedTabItemsGroups.ToList()) {
                if (group.Items.Contains(tabItem)) {
                    this.RemoveTabItemFromGroup(tabItem, group);
                    break;
                }
            }
        }


        private void RemoveTabItemFromGroup(TabItemBase tabItem, TabItemsGroupBase group) {
            if (group.Items.Remove(tabItem)) {
                Helpers.Diagnostic.Logger.LogDebug($"Removed tab \"{tabItem.Caption}\" from group \"{group.GroupName}\"");

                if (!group.Items.Any()) {
                    this.RemoveTabItemsGroup(group);
                }
            }
        }


        private void RemoveTabItemsGroup(TabItemsGroupBase tabItemsGroup) {
            if (this.SortedTabItemsGroups.Remove(tabItemsGroup)) {
                Helpers.Diagnostic.Logger.LogDebug($"Removed group \"{tabItemsGroup.GroupName}\"");
                this.UpdateSeparatorsBetweenGroups();
            }
        }


        private void UpdateSeparatorsBetweenGroups() {
            // Remove existing separators
            foreach (var sep in this.SortedTabItemsGroups.OfType<SeparatorTabItemsGroup>().ToList()) {
                this.SortedTabItemsGroups.Remove(sep);
            }

            if (this.HasGroup<TabItemsPreviewGroup>() &&
                (this.HasGroup<TabItemsPinnedGroup>() || this.HasGroup<TabItemsDefaultGroup>())) {
                this.SortedTabItemsGroups.Add(new SeparatorTabItemsGroup("Preview-Pinned"));
            }

            if (this.HasGroup<TabItemsPinnedGroup>() && this.HasGroup<TabItemsDefaultGroup>()) {
                this.SortedTabItemsGroups.Add(new SeparatorTabItemsGroup("Pinned-Default"));
            }
        }


        //
        // ░ История закрытых вкладок
        //
        private ClosedTabEntry CreateClosedTabEntry(TabItemBase tabItem) {
            // В историю не кладём COM-объекты EnvDTE: после закрытия они становятся недействительными.
            // Для повторного открытия достаточно стабильного пути/WindowId и описания группы.
            var current = this.FindTabItemWithGroup(tabItem);
            var group = current?.Group;

            return new ClosedTabEntry {
                Kind = tabItem is TabItemWindow ? ClosedTabKind.ToolWindow : ClosedTabKind.Document,
                FullName = tabItem.FullName,
                WindowId = (tabItem as TabItemWindow)?.WindowId,
                GroupName = group?.GroupName ?? string.Empty,
                GroupKind = this.GetClosedTabGroupKind(group)
            };
        }

        private void PushClosedTabsOperation(IEnumerable<ClosedTabEntry> entries) {
            var operationEntries = entries.ToList();
            if (operationEntries.Count == 0) {
                return;
            }

            _closedTabsHistory.Push(new ClosedTabsOperation(operationEntries));
            if (_closedTabsHistory.Count <= ClosedTabsHistoryCapacity) {
                return;
            }

            // Stack не умеет удалять самый старый элемент, поэтому пересобираем его без хвоста.
            var retainedOperations = _closedTabsHistory.Take(ClosedTabsHistoryCapacity).Reverse().ToList();
            _closedTabsHistory.Clear();
            foreach (var operation in retainedOperations) {
                _closedTabsHistory.Push(operation);
            }
        }

        private void RestoreLastClosedTabsOperation() {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_closedTabsHistory.Count == 0) {
                return;
            }

            var operation = _closedTabsHistory.Pop();
            // OpenFile/Show активируют каждый восстановленный frame и синхронно вызывают наши
            // DTE-обработчики. Флаг заставляет их игнорировать Ctrl от комбинации Ctrl+Z.
            _isRestoringClosedTabs = true;
            try {
                foreach (var entry in operation.Entries) {
                    try {
                        TabItemBase? restoredTabItem = entry.Kind == ClosedTabKind.Document
                            ? this.RestoreClosedDocument(entry)
                            : this.RestoreClosedToolWindow(entry);

                        if (restoredTabItem != null) {
                            this.MoveRestoredTabToOriginalGroup(restoredTabItem, entry);
                        }
                    }
                    catch (Exception ex) {
                        Helpers.Diagnostic.Logger.LogError($"Failed to restore closed tab '{entry.FullName}': {ex}");
                    }
                }
            }
            finally {
                _isRestoringClosedTabs = false;
            }

            // Контекстный Idle может наступить уже после следующего быстрого Ctrl+A. Сразу
            // возвращаем WPF-фокус панели, а отложенный запрос ниже страхует от поздних
            // DocumentOpened/activation-событий восстановленных VS-фреймов.
            this.FocusEditModeInputTarget();
            _keyboardTabNavigationExtension.RestoreInputTarget();
        }

        private TabItemDocument? RestoreClosedDocument(ClosedTabEntry entry) {
            // Повторный вызов безопасен: уже открытый документ не открываем ещё раз, а только
            // возвращаем существующий TabItem в сохранённую группу.
            var existingTabItem = this.FindTabItem(entry.FullName);
            if (existingTabItem != null) {
                return existingTabItem;
            }
            if (!File.Exists(entry.FullName)) {
                Helpers.Diagnostic.Logger.LogWarning($"Cannot restore deleted document '{entry.FullName}'");
                return null;
            }

            var window = PackageServices.Dte2.ItemOperations.OpenFile(entry.FullName);
            // Обычно OnDocumentOpened уже успевает добавить вкладку синхронно. Поиск по пути —
            // запасной вариант для видов редактора, у которых OpenFile не вернул Document.
            var restoredTabItem = window?.Document == null ? null : this.FindTabItem(window.Document);
            restoredTabItem ??= this.FindTabItem(entry.FullName);
            return restoredTabItem;
        }

        private TabItemWindow? RestoreClosedToolWindow(ClosedTabEntry entry) {
            // Persistence GUID позволяет заново создать tool window без хранения устаревшего
            // EnvDTE.Window из момента закрытия.
            if (!Guid.TryParse(entry.WindowId, out var persistenceGuid)) {
                return null;
            }

            var uiShell = Package.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
            if (uiShell == null) {
                return null;
            }

            int result = uiShell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fForceCreate, ref persistenceGuid, out var frame);
            if (ErrorHandler.Failed(result) || frame == null) {
                return null;
            }

            frame.Show();
            this.UpdateWindowTabsInfo();
            return this.SortedTabItemsGroups
                .SelectMany(group => group.Items)
                .OfType<TabItemWindow>()
                .FirstOrDefault(item => string.Equals(item.WindowId, entry.WindowId, StringComparison.OrdinalIgnoreCase));
        }

        private void MoveRestoredTabToOriginalGroup(TabItemBase tabItem, ClosedTabEntry entry) {
            // OnDocumentOpened сначала помещает документ в автоматически вычисленную группу.
            // После этого переносим тот же объект в группу, сохранённую перед закрытием.
            var current = this.FindTabItemWithGroup(tabItem);
            if (current != null) {
                this.RemoveTabItemFromGroup(current.Value.Item, current.Value.Group);
            }

            this.AddTabItemToGroupIfMissing(tabItem, this.CreateRestoredGroup(entry));
            if (tabItem is TabItemDocument tabItemDocument && entry.GroupKind == ClosedTabGroupKind.Pinned) {
                // Одного UI-флага недостаточно: закрепляем также настоящий frame Visual Studio.
                tabItemDocument.ShellDocument.OpenDocumentAsPinned();
            }
        }

        private TabItemsGroupBase CreateRestoredGroup(ClosedTabEntry entry) {
            return entry.GroupKind switch {
                ClosedTabGroupKind.Preview => new TabItemsPreviewGroup(),
                ClosedTabGroupKind.Pinned => new TabItemsPinnedGroup(entry.GroupName),
                _ => new TabItemsDefaultGroup(entry.GroupName)
            };
        }

        private ClosedTabGroupKind GetClosedTabGroupKind(TabItemsGroupBase? group) {
            if (group is TabItemsPreviewGroup) {
                return ClosedTabGroupKind.Preview;
            }
            if (group is TabItemsPinnedGroup) {
                return ClosedTabGroupKind.Pinned;
            }
            return ClosedTabGroupKind.Default;
        }

        private string GetHistoryKey(TabItemBase tabItem) {
            // Префикс исключает случайное совпадение пути документа с идентификатором окна.
            return tabItem is TabItemWindow tabItemWindow
                ? $"window:{tabItemWindow.WindowId}"
                : $"document:{tabItem.FullName}";
        }

        private void SelectActivatedTabItem(TabItemBase tabItem) {
            if (_isRestoringClosedTabs) {
                // KeyUp для Ctrl+Z не может обработаться, пока синхронное восстановление не завершено.
                // Явно задаём обычный одиночный выбор и не читаем удерживаемый Keyboard.Modifiers.
                _tabNavigationController.SetSelectionWithoutActivation(tabItem, true, ModifierKeys.None);
                return;
            }

            tabItem.IsSelected = true;
        }

        private bool HasGroup<T>() where T : TabItemsGroupBase {
            return this.SortedTabItemsGroups.OfType<T>().Any();
        }


        // 
        // ░ Updating tabs
        //
        // Обновление документа в UI после изменения или переименования
        private void UpdateDocumentUI(string oldPath, string newPath = null) {
            foreach (var group in this.SortedTabItemsGroups) {
                var docInfo = group.Items.FirstOrDefault(d => string.Equals(d.FullName, oldPath, StringComparison.OrdinalIgnoreCase));
                if (docInfo != null) {
                    if (newPath == null) {
                        // Обновляем только имя (в случае изменения)
                        docInfo.Caption = Path.GetFileName(oldPath);
                    }
                    else {
                        // Обновляем имя и полный путь (в случае переименования)
                        docInfo.FullName = newPath;
                        docInfo.Caption = Path.GetFileName(newPath);
                    }
                    return;
                }
            }
        }

        private void UpdateWindowTabsInfo() {
            ThreadHelper.ThrowIfNotOnUIThread();

            // DTE.Windows — COM-коллекция, поэтому перечисляем её один раз и строим индекс.
            // Иначе для каждой виртуальной вкладки пришлось бы заново обходить все окна Visual Studio.
            var windowsById = PackageServices.Dte2.Windows
                .Cast<EnvDTE.Window>()
                .Select(window => new {
                    Window = window,
                    Id = VsShell.Document.ShellWindow.GetWindowId(window)
                })
                .Where(entry => !string.IsNullOrEmpty(entry.Id))
                .GroupBy(entry => entry.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Window, StringComparer.OrdinalIgnoreCase);

            this.ForEachTab<TabItemWindow>(tabItemWindow => {
                try {
                    // По стабильному WindowId находим актуальное окно и переносим изменившийся заголовок в UI.
                    if (windowsById.TryGetValue(tabItemWindow.WindowId, out var matchingWindow) && tabItemWindow.Caption != matchingWindow.Caption) {
                        Helpers.Diagnostic.Logger.LogDebug($"Updating TabItemWindow caption: '{tabItemWindow.Caption}' → '{matchingWindow.Caption}'");
                        tabItemWindow.Caption = matchingWindow.Caption;
                        tabItemWindow.FullName = matchingWindow.Caption;
                    }
                }
                catch (Exception ex) {
                    Helpers.Diagnostic.Logger.LogError($"Failed to update caption for TabItemWindow: {ex.Message}");
                }
            });
        }


        // 
        // ░ Activating tabs
        //
        private void ActivatePrimaryTabItem() {
            //using var __logFunctionScoped = Helpers.Diagnostic.Logger.LogFunctionScope("ActivatePrimaryTabItem()");
            ThreadHelper.ThrowIfNotOnUIThread();

            var primaryTabItem = _tabItemsSelectionCoordinator.PrimarySelection?.Item;
            if (primaryTabItem is IActivatableTab activatableTab) {
                Helpers.Diagnostic.Logger.LogDebug($"Activate - \"{primaryTabItem.Caption}\"");
                activatableTab.Activate();
            }
        }

        // TODO:

        internal void RegisterTabItemControl(TabItemControl control) {
            _tabItemControls.Add(control);
            this.ApplyScaleToTabItemControl(control);

            if (ReferenceEquals(control.DataContext, _keyboardTabNavigationExtension?.FocusedItem)) {
                control.IsEditFocused = true;
            }
        }

        internal void UnregisterTabItemControl(TabItemControl control) {
            _tabItemControls.Remove(control);
        }

        internal bool HandleTabEditKey(Key key, ModifierKeys modifiers) {
            if (!this.IsTabEditMode) {
                return false;
            }

            bool isControlPressed = (modifiers & ModifierKeys.Control) == ModifierKeys.Control;
            if (key == Key.F2 && modifiers == ModifierKeys.None) {
                var focusedControl = this.FindTabItemControl(_keyboardTabNavigationExtension.FocusedItem);
                if (focusedControl != null) {
                    focusedControl.BeginRename();
                }
                return true;
            }

            if (key == Key.Escape && modifiers == ModifierKeys.None) {
                // Escape сворачивает мультивыбор к пунктирно сфокусированной вкладке. Так панель
                // выходит из группового действия, но сохраняет понятную точку навигации.
                var focusedItem = _keyboardTabNavigationExtension.FocusedItem;
                if (focusedItem != null) {
                    _tabItemsSelectionCoordinator.SetSelection(focusedItem, true, ModifierKeys.None);
                    _keyboardTabNavigationExtension.RestoreInputTarget();
                }

                return true;
            }

            if (isControlPressed && key == Key.A) {
                Helpers.Diagnostic.Logger.LogDebug("[NavigationInput] Ctrl+A handled by tabs panel: selecting all tabs.");
                // SelectAll использует coordinator, чтобы корректно обновить общий selection state
                // и все визуальные состояния групп, а не только IsSelected отдельных моделей.
                _tabNavigationController.SelectAll();
                _keyboardTabNavigationExtension.RestoreInputTarget();
                return true;
            }

            if (key == Key.Delete && modifiers == ModifierKeys.None) {
                var focusedItem = _keyboardTabNavigationExtension.FocusedItem;
                if (focusedItem == null) {
                    return true;
                }

                var selectedItems = _tabItemsSelectionCoordinator.SelectedItems;
                // Если навигационный фокус входит в мультивыбор, Delete относится ко всей пачке.
                // Иначе закрывается только пунктирно сфокусированная вкладка.
                var itemsToClose = focusedItem.IsSelected && selectedItems.Count > 1
                    ? selectedItems.Select(entry => entry.Item).ToList()
                    : new List<TabItemBase> { focusedItem };
                this.CloseTabItems(itemsToClose);

                var nextFocusedItem = this.GetVisibleTabItems().FirstOrDefault();
                if (nextFocusedItem != null) {
                    // Закрытие активного frame может вернуть command target редактору. Повторно
                    // закрепляем навигационный фокус на оставшейся вкладке панели.
                    _keyboardTabNavigationExtension.FocusItem(nextFocusedItem);
                    _keyboardTabNavigationExtension.RestoreInputTarget();
                }
                return true;
            }

            if (isControlPressed && key == Key.Z) {
                Helpers.Diagnostic.Logger.LogDebug($"[NavigationInput] Ctrl+Z received by tabs panel. HistoryCount={_closedTabsHistory.Count}.");
                // Пустая история не должна поглощать стандартный Undo редактора.
                if (_closedTabsHistory.Count == 0) {
                    return false;
                }

                this.RestoreLastClosedTabsOperation();
                return true;
            }

            return _tabNavigationController.HandleKey(key, modifiers);
        }

        internal bool TryRenameTabItem(TabItemControl source, string proposedName) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (source.DataContext is not TabItemDocument tabItemDocument) {
                this.ShowTabRenameMessage("Only document tabs can be renamed.");
                return false;
            }

            string oldPath = tabItemDocument.FullName;
            string oldExtension = Path.GetExtension(oldPath);
            string newName = proposedName?.Trim() ?? string.Empty;
            // F2 может завершиться без фактического редактирования. Проверяем это до ограничений
            // расширения, чтобы Escape через потерю фокуса и обычный Enter не показывали ошибку.
            if (string.Equals(Path.GetFileName(oldPath), newName, StringComparison.Ordinal)) {
                return true;
            }

            if (!RenamableTabExtensions.Contains(oldExtension)) {
                this.ShowTabRenameMessage($"Files with the '{oldExtension}' extension cannot be renamed from Tabs Manager.");
                return false;
            }
            if (string.IsNullOrWhiteSpace(newName) ||
                !string.Equals(newName, Path.GetFileName(newName), StringComparison.Ordinal) ||
                newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {

                this.ShowTabRenameMessage("Enter a valid file name without a directory path.");
                return false;
            }

            string newExtension = Path.GetExtension(newName);
            if (!RenamableTabExtensions.Contains(newExtension)) {
                this.ShowTabRenameMessage($"The target extension '{newExtension}' is not supported. Use .c, .cpp, .h, .cxx, or .hxx.");
                return false;
            }

            string? directory = Path.GetDirectoryName(oldPath);
            string newPath = directory == null ? newName : Path.Combine(directory, newName);
            // На Windows case-only rename указывает на тот же файл и не является конфликтом.
            if (!string.Equals(oldPath, newPath, StringComparison.OrdinalIgnoreCase) && File.Exists(newPath)) {
                this.ShowTabRenameMessage($"A file named '{newName}' already exists in this directory.");
                return false;
            }

            EnvDTE.ProjectItem? projectItem;
            try {
                projectItem = tabItemDocument.ShellDocument.Document.ProjectItem;
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to resolve ProjectItem for '{oldPath}': {ex}");
                this.ShowTabRenameMessage("Visual Studio could not resolve the project item for this document.");
                return false;
            }

            if (projectItem == null) {
                this.ShowTabRenameMessage("This document is not represented by a project item and cannot be safely renamed.");
                return false;
            }

            try {
                // ProjectItem.Name использует штатную project-system рутину: она обновляет файл,
                // открытый document moniker и ссылки проекта так же, как Solution Explorer.
                projectItem.Name = newName;
                tabItemDocument.Caption = tabItemDocument.ShellDocument.Document.Name;
                tabItemDocument.FullName = tabItemDocument.ShellDocument.Document.FullName;
                return true;
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to rename tab '{oldPath}' to '{newName}': {ex}");
                this.ShowTabRenameMessage($"Visual Studio could not rename '{Path.GetFileName(oldPath)}' to '{newName}'.\n\n{ex.Message}");
                return false;
            }
        }

        internal void RestoreTabNavigationInputTarget() {
            if (this.IsTabEditMode) {
                _keyboardTabNavigationExtension.RestoreInputTarget();
            }
        }

        private void ShowTabRenameMessage(string message) {
            ThreadHelper.ThrowIfNotOnUIThread();
            VsShellUtilities.ShowMessageBox(
                ServiceProvider.GlobalProvider,
                message,
                "Rename Tab",
                OLEMSGICON.OLEMSGICON_WARNING,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST
            );
        }

        private bool CanHandleRedirectedNavigationKey(Key key) {
            // Visual Studio разрешает глобальные команды до WPF PreviewKeyDown. Поэтому OLE-фильтр
            // спрашивает контрол заранее, кому принадлежит команда: панели или редактору.
            if (!this.IsKeyboardFocusWithin) {
                return false;
            }

            // При inline rename фильтр должен поглотить OLE-команду редактора. Сервис ввода
            // перенаправит её в реально сфокусированный TextBox, не затрагивая editor view.
            if (Keyboard.FocusedElement is DependencyObject focusedElement &&
                Helpers.VisualTree.FindParentByType<TabItemControl>(focusedElement)?.IsRenaming == true) {
                return true;
            }

            // Ctrl+Z принадлежит панели только при наличии реальной операции восстановления.
            return key != Key.Z || _closedTabsHistory.Count > 0;
        }

        internal void HandleTabPointerNavigation(TabItemControl source, ModifierKeys modifiers) {
            if (source.DataContext is not TabItemBase tabItem) {
                return;
            }

            // Контроллер одинаково применяет правила выбора в обоих режимах: простой ЛКМ может
            // активировать вкладку, а Ctrl/Shift меняют набор выделения без обязательной активации.
            _tabNavigationController.OnPointerSelection(tabItem, modifiers);

            if (this.IsTabEditMode) {
                // Мышь также переносит пунктирный навигационный фокус, чтобы следующая стрелка
                // продолжила движение от вкладки, на которую только что нажал пользователь.
                _keyboardTabNavigationExtension.FocusItem(tabItem);

                // Активация документа в результате клика может вернуть command target редактору
                // уже после MouseUp. Планируем восстановление даже при клике по той же вкладке.
                _keyboardTabNavigationExtension.RestoreInputTarget();
            }
        }

        private List<TabItemBase> GetVisibleTabItems() {
            return this.SortedTabItemsGroups
                .SelectMany(group => group.Items)
                .ToList();
        }

        private void OnKeyboardFocusedTabItemChanged(TabItemBase? tabItem) {
            var previousControl = _tabItemControls.FirstOrDefault(control => control.IsEditFocused);
            if (previousControl != null) {
                previousControl.IsEditFocused = false;
            }

            var control = this.FindTabItemControl(tabItem);
            if (control == null) {
                return;
            }

            control.IsEditFocused = true;
            control.BringIntoView();
            this.FocusEditModeInputTarget();
        }

        private void OnKeyboardInputTargetRestoreRequested() {
            this.Dispatcher.BeginInvoke(
                new Action(this.FocusEditModeInputTarget),
                DispatcherPriority.ContextIdle
            );
        }

        private TabItemControl? FindTabItemControl(TabItemBase? tabItem) {
            return tabItem == null
                ? null
                : _tabItemControls.FirstOrDefault(candidate => ReferenceEquals(candidate.DataContext, tabItem));
        }

        private void FocusEditModeInputTarget() {
            // Контейнеры ListView намеренно нефокусируемые, чтобы обычная работа со вкладками
            // не отбирала command target у редактора. В edit mode используем существующий
            // FocusStealer как единую WPF-цель ввода для корневого PreviewKeyDown.
            FocusManager.SetFocusedElement(this, this.FocusStealer);
            this.FocusStealer.Focus();
            Keyboard.Focus(this.FocusStealer);
        }

        private void ApplyScaleToTabItemControl(TabItemControl control) {
            // Локальное значение имеет приоритет над дефолтом из подключённого BrushResources.xaml.
            control.Resources["AppTabItemHeight"] = this.ScaleFactorTabsCompactness * 26.4;
        }

        private void SyncActiveDocumentWithPrimaryTabItem() {
            ThreadHelper.ThrowIfNotOnUIThread();

            var activeWindow = PackageServices.Dte2.ActiveWindow;
            if (activeWindow == null) {
                return;
            }

            // Сначала синхронизируем именно активный VS-фрейм. Это состояние независимо от
            // PrimarySelection и гарантирует, что фиолетовая рамка не следует за мультивыбором.
            TabItemBase activeFrameTabItem = null;
            if (VsShell.Document.ShellWindow.IsTabWindow(activeWindow)) {
                if (activeWindow.Document != null) {
                    activeFrameTabItem = this.FindTabItem(activeWindow.Document);
                }
                else {
                    activeFrameTabItem = this.FindTabItem(activeWindow);
                }
            }
            else if (PackageServices.Dte2.ActiveDocument != null) {
                activeFrameTabItem = this.FindTabItem(PackageServices.Dte2.ActiveDocument);
            }

            if (activeFrameTabItem != null) {
                this.SetActiveFrameTabItem(activeFrameTabItem);
            }

            var selectedTabItem = _tabItemsSelectionCoordinator.PrimarySelection?.Item;
            TabItemBase targetTabItem = null;

            // Затем отдельно синхронизируем выбор. Этот путь нужен для активаций, пришедших
            // извне панели (например, из Solution Explorer), где pointer-навигация не участвовала.
            if (VsShell.Document.ShellWindow.IsTabWindow(activeWindow)) {
                // В области вкладок VS может быть активен как документ, так и tool window.
                if (activeWindow.Document == null) {
                    if (string.Equals(activeWindow.Caption, selectedTabItem?.Caption, StringComparison.OrdinalIgnoreCase)) {
                        return;
                    }
                    Helpers.Diagnostic.Logger.LogDebug($"Sync tabs with activeWindow.Caption = {activeWindow.Caption}");
                    targetTabItem = this.FindTabItem(activeWindow);
                }
                else {
                    if (string.Equals(activeWindow.Document.FullName, selectedTabItem?.FullName, StringComparison.OrdinalIgnoreCase)) {
                        return;
                    }
                    Helpers.Diagnostic.Logger.LogDebug($"Sync tabs with activeWindow.Document.Name = {activeWindow.Document.Name}");
                    targetTabItem = this.FindTabItem(activeWindow.Document);
                }
            }
            else {
                // Если ActiveWindow не является вкладкой, ориентируемся на ActiveDocument:
                // например, документ мог быть выбран через Solution Explorer.
                var activeDocument = PackageServices.Dte2.ActiveDocument;
                if (activeDocument == null) {
                    return;
                }
                if (string.Equals(activeDocument.FullName, selectedTabItem?.FullName, StringComparison.OrdinalIgnoreCase)) {
                    return;
                }
                Helpers.Diagnostic.Logger.LogDebug($"Sync tabs with activeDocument.Name = {activeDocument.Name}");
                targetTabItem = this.FindTabItem(activeDocument);
            }

            if (targetTabItem != null) {
                // Меняем только выбор; актуальная фиолетовая рамка уже выставлена выше
                // по фактическому активному фрейму и от PrimarySelection не зависит.
                this.SelectActivatedTabItem(targetTabItem);
            }
        }

        private void SetActiveFrameTabItem(TabItemBase activeTabItem) {
            // В каждый момент только одна представленная вкладка может соответствовать
            // активному VS-фрейму; мультивыбор на этот marker не влияет.
            foreach (var tabItem in this.SortedTabItemsGroups.SelectMany(group => group.Items)) {
                tabItem.Metadata.SetFlag("IsFrameActive", ReferenceEquals(tabItem, activeTabItem));
            }
        }


        //
        // ░ Helpers
        // ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 
        //
        private TabItemDocument? FindTabItem(EnvDTE.Document document) {
            return this.FindTabItemWithGroup(document)?.Item;
        }

        private TabItemWindow? FindTabItem(EnvDTE.Window window) {
            return this.FindTabItemWithGroup(window)?.Item;
        }

        private TabItemBase? FindTabItem(TabItemBase tabItem) {
            return this.FindTabItemWithGroup(tabItem)?.Item;
        }

        private TabItemDocument? FindTabItem(string documentFullName) {
            return this.FindTabItemWithGroup(documentFullName)?.Item;
        }


        private (TabItemDocument Item, TabItemsGroupBase Group)? FindTabItemWithGroup(EnvDTE.Document document) {
            ThreadHelper.ThrowIfNotOnUIThread();
            return this.FindTabItemWithGroup(document.FullName);
        }

        private (TabItemWindow Item, TabItemsGroupBase Group)? FindTabItemWithGroup(EnvDTE.Window window) {
            var result = this.FindTabItemWithGroup(new TabItemWindow(window));
            if (result is { Item: TabItemWindow win, Group: var group }) {
                return (win, group);
            }
            return null;
        }

        private (TabItemBase Item, TabItemsGroupBase Group)? FindTabItemWithGroup(TabItemBase tabItem) {
            if (tabItem is TabItemWindow tabItemWindow) {
                return this.FindTabItemWithGroupBy<TabItemWindow>(
                    w => string.Equals(w.WindowId, tabItemWindow.WindowId, StringComparison.OrdinalIgnoreCase));
            }
            return this.FindTabItemWithGroupBy<TabItemBase>(
                t => string.Equals(t.FullName, tabItem.FullName, StringComparison.OrdinalIgnoreCase));
        }

        private (TabItemDocument Item, TabItemsGroupBase Group)? FindTabItemWithGroup(string documentFullName) {
            return this.FindTabItemWithGroupBy<TabItemDocument>(
                    d => string.Equals(d.FullName, documentFullName, StringComparison.OrdinalIgnoreCase));
        }


        private (T Item, TabItemsGroupBase Group)? FindTabItemWithGroupBy<T>(Func<T, bool> predicate) where T : TabItemBase {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (var group in this.SortedTabItemsGroups) {
                var match = group.Items
                    .OfType<T>()
                    .FirstOrDefault(predicate);

                if (match != null) {
                    return (match, group);
                }
            }

            return null;
        }


        private void ForEachTab<T>(Action<T> action) where T : TabItemBase {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (var group in this.SortedTabItemsGroups) {
                foreach (var tabItem in group.Items.OfType<T>()) {
                    action(tabItem);
                }
            }
        }

        private bool IsTemporaryFile(string fullPath) {
            string extension = Path.GetExtension(fullPath);
            return extension.Equals(".TMP", StringComparison.OrdinalIgnoreCase) ||
                   fullPath.Contains("~") && fullPath.Contains(".TMP");
        }

        private enum ClosedTabKind {
            Document,
            ToolWindow
        }

        private enum ClosedTabGroupKind {
            Default,
            Pinned,
            Preview
        }

        private sealed class ClosedTabEntry {
            // Запись содержит только устойчивые данные, пригодные после уничтожения VS frame.
            public ClosedTabKind Kind { get; set; }
            public string FullName { get; set; } = string.Empty;
            public string? WindowId { get; set; }
            public string GroupName { get; set; } = string.Empty;
            public ClosedTabGroupKind GroupKind { get; set; }
        }

        private sealed class ClosedTabsOperation {
            // Entries сохраняет границу пользовательской операции для атомарного Ctrl+Z.
            public IReadOnlyList<ClosedTabEntry> Entries { get; }

            public ClosedTabsOperation(IReadOnlyList<ClosedTabEntry> entries) {
                this.Entries = entries;
            }
        }
    }
}
