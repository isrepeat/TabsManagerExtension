using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using Helpers.Ex;
using TMEx = TabsManagerExtension;

namespace TabsManagerExtension.Controls {
    public partial class TabsManagerToolWindowControl : Helpers.BaseUserControl {

        // Properties:
        private Helpers.Collections.SortedObservableCollection<TMEx.State.Document.TabItemsGroupBase> _sortedTabItemsGroups;
        public Helpers.Collections.SortedObservableCollection<TMEx.State.Document.TabItemsGroupBase> SortedTabItemsGroups {
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
                    Settings.TabsManagerSettingsService.SetTabsScaleFactor(value);
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
                    Settings.TabsManagerSettingsService.SetTabEditMode(value);

                    if (value) {
                        this.UpdateEditModeInputRedirect();
                        VsixThreadHelper.RunOnUiThread(
                            this.Dispatcher,
                            () => _keyboardTabNavigationExtension?.InitializeFocus(),
                            DispatcherPriority.Normal
                        );
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
        private TMEx.Controls.Tabs.VisualStudioTabEventBridge? _visualStudioTabEventBridge;

        private TMEx.Controls.Tabs.TabsStateReconciler? _tabsStateReconciler;
        private TMEx.Controls.Tabs.ToolWindowSessionManager? _toolWindowSessionManager;
        private readonly TMEx.Controls.Tabs.SolutionFileChangeMonitor _solutionFileChangeMonitor;
        private TMEx.Controls.Tabs.TabCollectionManager _tabCollectionManager;
        private readonly TMEx.Controls.Tabs.ClosedTabsHistory _closedTabsHistory = new();
        private TMEx.Controls.Tabs.ClosedTabsRestorer? _closedTabsRestorer;

        private Helpers.Collections.GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase> _tabItemsSelectionCoordinator;
        private Navigation.TabNavigationController _tabNavigationController;
        private Navigation.KeyboardTabNavigationExtension _keyboardTabNavigationExtension;
        private VsShell.TextEditor.Overlay.TextEditorOverlayController _textEditorOverlayController;
        private readonly TMEx.Controls.Tabs.TabAppearanceManager _tabAppearanceManager = new();
        private readonly TMEx.Controls.Tabs.TabRenameService _tabRenameService = new();
        private readonly TMEx.Controls.Tabs.TabMenuItemFactory _tabMenuItemFactory;
        private TMEx.Controls.Tabs.ProjectContextController? _projectContextController;
        private TMEx.Controls.Tabs.VirtualMenuController? _virtualMenuController;
        private TMEx.Controls.Tabs.TabActivationSynchronizer? _tabActivationSynchronizer;
        private TMEx.Controls.Tabs.TabsWorkspaceSynchronizer? _tabsWorkspaceSynchronizer;
        private TMEx.Controls.Tabs.TabInputController? _tabInputController;
        private TMEx.Controls.Tabs.TabCommandController? _tabCommandController;
        // Текущий UX: только обычный ЛКМ меняет активный документ; Ctrl/Shift/Space
        // управляют selection, не перемещая фиолетовую рамку активного VS-фрейма.
        private const Navigation.TabSelectionActivationPolicy TabSelectionActivationPolicy =
            Navigation.TabSelectionActivationPolicy.ActivateOnlyOnUnmodifiedPointerSelection;
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
        public ICommand OnCopySelectedTabNamesCommand { get; }

        public TabsManagerToolWindowControl() {
            this.InitializeComponent();
            _solutionFileChangeMonitor = new TMEx.Controls.Tabs.SolutionFileChangeMonitor(this.Dispatcher);
            _solutionFileChangeMonitor.FileChanged += this.OnFileChanged;
            _solutionFileChangeMonitor.FileRenamed += this.OnFileRenamed;
            _solutionFileChangeMonitor.FileDeleted += this.OnFileDeleted;
            this.ScaleFactorTabsCompactness = Settings.TabsManagerSettingsService.TabsScaleFactor;
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
            this.OnCopySelectedTabNamesCommand = new Helpers.RelayCommand<object>(this.OnCopyTabName);

            _tabMenuItemFactory = new TMEx.Controls.Tabs.TabMenuItemFactory(
                this.OnPinTabItem,
                this.OnCopyTabName,
                this.OnCopyTabPath,
                this.OnOpenLocationTabItem,
                this.OnCloseTabItem,
                this.OnCloseSelectedTabItems,
                this.OnMoveTabItemToRelatedProject,
                this.OnReloadDocumentReferencesProjects
            );

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
                new Helpers.MenuItemHeader {
                    Header = $"TabsManagerExtension v{TabsManagerExtensionPackage.GetInstalledExtensionVersion()}"
                },
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
            Settings.TabsManagerSettingsService.TabsScaleFactorChanged += this.OnTabsScaleFactorChanged;
            Settings.TabsManagerSettingsService.AppearanceChanged += this.OnAppearanceChanged;
            VsShell.TextEditor.Services.TextEditorInputCommandFilterService.Instance.AddTrackedInputElement(this);
            // Режим восстанавливаем только после регистрации: его включение перенаправляет
            // команды редактора в эту панель через TextEditorInputCommandFilterService.
            this.IsTabEditMode = Settings.TabsManagerSettingsService.IsTabEditMode;
            this.UpdateEditModeInputRedirect();
            Services.ExtensionStatusService.Instance.FeatureReadinessChanged += this.OnFeatureReadinessChanged;
            this.IsIncludeGraphReady = Services.ExtensionStatusService.Instance.IsFeatureReady(
                Services.ExtensionStatusService.IncludeGraphFeature
            );

            this.InitializeDTE();
            this.InitializeFileWatcher();
            this.InitializeTabItemsSelectionCoordinator();
            this.InitializeBackgroundRoutine();
            _toolWindowSessionManager?.PrepareActiveWindowRestore();
            this.ApplyScaleTabsCompactness();
            this.ApplyAppearance();

            var hierarchyAnalyzer = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance;
            hierarchyAnalyzer.InitialAnalysisCompleted.Add(this.OnInitialHierarchyAnalysisCompleted);
            hierarchyAnalyzer.InitialAnalysisCompleted.InvokeForLastHandlerIfTriggered();
            var solutionEvents = VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance;
            solutionEvents.SolutionClosing.Add(this.OnSolutionClosing);
            solutionEvents.SolutionClosed.Add(this.OnSolutionClosed);
        }

        private void OnUnloaded(object sender, RoutedEventArgs e) {
            Settings.TabsManagerSettingsService.SetTabEditMode(this.IsTabEditMode);
            Settings.TabsManagerSettingsService.TabsScaleFactorChanged -= this.OnTabsScaleFactorChanged;
            Settings.TabsManagerSettingsService.AppearanceChanged -= this.OnAppearanceChanged;
            VsShell.TextEditor.Services.TextEditorInputCommandFilterService.Instance.SetForcedInputTarget(null);
            VsShell.TextEditor.Services.TextEditorInputCommandFilterService.Instance.RemoveTrackedInputElement(this);
            var solutionEvents = VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance;
            solutionEvents.SolutionClosed.Remove(this.OnSolutionClosed);
            solutionEvents.SolutionClosing.Remove(this.OnSolutionClosing);
            VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance.InitialAnalysisCompleted.Remove(this.OnInitialHierarchyAnalysisCompleted);
            Services.ExtensionStatusService.Instance.FeatureReadinessChanged -= this.OnFeatureReadinessChanged;
            _toolWindowSessionManager?.Save(VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.IsSolutionClosing);
            _toolWindowSessionManager?.CancelActiveWindowRestore();
            this.UninitializeTabItemsSelectionCoordinator();
            this.UninitializeFileWatcher();
            this.UninitializeBackgroundRoutine();
            this.UninitializeDTE();

            Services.ExtensionServices.EndUsage();
        }

        private void OnTabsScaleFactorChanged(double value) {
            VsixThreadHelper.RunOnUiThread(
                this.Dispatcher,
                () => this.ScaleFactorTabsCompactness = value,
                DispatcherPriority.Normal
            );
        }

        private void OnAppearanceChanged() {
            if (this.Dispatcher.CheckAccess()) {
                this.ApplyAppearance();
                return;
            }

            VsixThreadHelper.RunOnUiThread(
                this.Dispatcher,
                this.ApplyAppearance,
                DispatcherPriority.Normal
            );
        }

        private void OnInitialHierarchyAnalysisCompleted(string solutionName) {
            _tabsWorkspaceSynchronizer?.EnsureSolutionLoaded(solutionName, () => {
                if (!_solutionFileChangeMonitor.IsRunning) {
                    this.InitializeFileWatcher();
                }
            });
        }

        private void OnFeatureReadinessChanged(string feature, bool isReady) {
            if (feature == Services.ExtensionStatusService.IncludeGraphFeature) {
                this.IsIncludeGraphReady = isReady;
            }
        }

        private void OnPreviewMouseDown(object sender, MouseButtonEventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Правая кнопка открывает контекстное меню и никогда не меняет текущий набор
            // выбранных вкладок. Это также исключает сброс selection при повторном открытии Popup.
            if (e.ChangedButton != MouseButton.Left) {
                return;
            }

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

            // Popup контекстного меню находится в отдельном визуальном дереве. Его клик не
            // должен попасть в обработчик пустой области и свернуть мультивыбор вкладок.
            bool isMenuInteraction = originalSource is Controls.MenuControl ||
                originalSource != null && Helpers.VisualTree.FindParentByType<Controls.MenuControl>(originalSource) != null;

            // Popup может визуально находиться над DocumentContainer, хотя логически относится
            // к ComboBox в нижней панели. Общая проверка интерактивного пути существовала до
            // 5ac1151 и не даёт принять ComboBoxItem/TextBox за клик по пустой области вкладок.
            bool isInteractiveInteraction = originalSource != null && this.ex_HasInteractiveElementOnPathFrom(originalSource);
            // Клик по вкладке или любому интерактивному контролу имеет собственную семантику
            // и не считается кликом по пустой области.
            if (isInsideDocumentContainer && !isTabInteraction && !isButtonInteraction && !isMenuInteraction && !isInteractiveInteraction) {
                // Сбрасываем мультивыбор явно: повторная активация уже открытого документа
                // может не породить событие DTE после команд контекстного меню.
                var activeFrameTabItem = _tabCollectionManager.AllTabs
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
                    if (retainedTabItem is TMEx.State.Document.IActivatableTab activatableTab) {
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
            _toolWindowSessionManager = new TMEx.Controls.Tabs.ToolWindowSessionManager(PackageServices.Dte2, this.Dispatcher);
            _visualStudioTabEventBridge = new TMEx.Controls.Tabs.VisualStudioTabEventBridge(PackageServices.Dte2);
            _visualStudioTabEventBridge.DocumentOpened += this.OnDocumentOpened;
            _visualStudioTabEventBridge.DocumentSaved += this.OnDocumentSaved;
            _visualStudioTabEventBridge.DocumentClosing += this.OnDocumentClosing;
            _visualStudioTabEventBridge.WindowActivated += this.OnWindowActivated;
            _visualStudioTabEventBridge.WindowClosing += this.OnWindowClosing;
            _visualStudioTabEventBridge.DocumentActivatedExternally += this.OnDocumentActivatedExternally;
            _visualStudioTabEventBridge.WindowFrameActivated += this.OnVsWindowFrameActivated;
            _visualStudioTabEventBridge.Start();
        }

        private void UninitializeDTE() {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_visualStudioTabEventBridge == null) {
                return;
            }

            _visualStudioTabEventBridge.Stop();
            _visualStudioTabEventBridge.WindowFrameActivated -= this.OnVsWindowFrameActivated;
            _visualStudioTabEventBridge.DocumentActivatedExternally -= this.OnDocumentActivatedExternally;
            _visualStudioTabEventBridge.WindowClosing -= this.OnWindowClosing;
            _visualStudioTabEventBridge.WindowActivated -= this.OnWindowActivated;
            _visualStudioTabEventBridge.DocumentClosing -= this.OnDocumentClosing;
            _visualStudioTabEventBridge.DocumentSaved -= this.OnDocumentSaved;
            _visualStudioTabEventBridge.DocumentOpened -= this.OnDocumentOpened;
            _visualStudioTabEventBridge = null;
            _toolWindowSessionManager?.Dispose();
            _toolWindowSessionManager = null;
        }


        //
        // ░ FileWatcher 
        //
        private void InitializeFileWatcher() {
            ThreadHelper.ThrowIfNotOnUIThread();
            _solutionFileChangeMonitor.Start(PackageServices.Dte2.Solution.FullName);
        }

        private void UninitializeFileWatcher() {
            _solutionFileChangeMonitor.Stop();
        }


        //
        // ░ TabItemsSelectionCoordinator 
        //
        private void InitializeTabItemsSelectionCoordinator() {
            // Collection manager является единственным владельцем структуры групп. Публичное
            // WPF-свойство ссылается на ту же observable collection и не хранит отдельную копию.
            _tabCollectionManager = new TMEx.Controls.Tabs.TabCollectionManager();
            this.SortedTabItemsGroups = _tabCollectionManager.Groups;

            // Selection coordinator строится поверх уже опубликованных групп и сообщает control
            // только изменения, которые должны отражаться в binding и связанных командах.
            _tabItemsSelectionCoordinator = new Helpers.Collections.GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase>(this.SortedTabItemsGroups);
            _tabItemsSelectionCoordinator.OnItemSelectionChanged = this.OnTabItemSelectionChanged;
            _tabItemsSelectionCoordinator.OnSelectionStateChanged = this.OnSelectionStateChanged;

            _tabNavigationController = new Navigation.TabNavigationController(
                _tabItemsSelectionCoordinator,
                this.GetVisibleTabItems
            ) {
                SelectionActivationPolicy = TabSelectionActivationPolicy
            };
            // Keyboard extension хранит отдельный navigation focus. Он не равен selection при
            // мультивыборе и подключается к общему navigation controller как расширение.
            _keyboardTabNavigationExtension = new Navigation.KeyboardTabNavigationExtension(_tabNavigationController);
            _keyboardTabNavigationExtension.FocusedItemChanged += this.OnKeyboardFocusedTabItemChanged;
            _keyboardTabNavigationExtension.InputTargetRestoreRequested += this.OnKeyboardInputTargetRestoreRequested;
            _tabNavigationController.AddExtension(_keyboardTabNavigationExtension);

            // Следующий слой зависит от collection/selection/navigation, поэтому создаётся после
            // них. Контроллеры получают общие экземпляры и не формируют параллельное состояние.
            _textEditorOverlayController = new VsShell.TextEditor.Overlay.TextEditorOverlayController(PackageServices.Dte2);
            _projectContextController = new TMEx.Controls.Tabs.ProjectContextController(
                PackageServices.Dte2,
                this.VirtualMenuControl,
                _tabCollectionManager,
                _tabItemsSelectionCoordinator
            );
            _virtualMenuController = new TMEx.Controls.Tabs.VirtualMenuController(
                this.VirtualMenuControl,
                this.VirtualMenuItems,
                _tabMenuItemFactory,
                _projectContextController
            );
            // Restorer создаётся до activation synchronizer: последний читает его IsRestoring,
            // чтобы не активировать каждый документ во время пакетного восстановления.
            _closedTabsRestorer = new TMEx.Controls.Tabs.ClosedTabsRestorer(
                PackageServices.Dte2,
                _tabCollectionManager,
                this.UpdateWindowTabsInfo,
                _keyboardTabNavigationExtension.RestoreInputTarget,
                this.FocusEditModeInputTarget
            );
            _tabCommandController = new TMEx.Controls.Tabs.TabCommandController(
                this.VirtualMenuControl,
                _tabCollectionManager,
                _closedTabsHistory,
                _tabItemsSelectionCoordinator,
                this.CreateClosedTabEntry
            );
            // Input controller делегирует фактическое закрытие и восстановление composition root,
            // где уже объединены команды, history и управление WPF-фокусом.
            _tabInputController = new TMEx.Controls.Tabs.TabInputController(
                this,
                this.FocusStealer,
                this.Resources,
                this.Dispatcher,
                () => this.IsTabEditMode,
                () => this.ScaleFactorTabsCompactness,
                _tabCollectionManager,
                _tabItemsSelectionCoordinator,
                _tabNavigationController,
                _keyboardTabNavigationExtension,
                _closedTabsHistory,
                _tabAppearanceManager,
                this.CloseTabItems,
                this.RestoreLastClosedTabsOperation,
                () => this.OnCopyTabName(null)
            );
            _tabActivationSynchronizer = new TMEx.Controls.Tabs.TabActivationSynchronizer(
                PackageServices.Dte2,
                this.Dispatcher,
                _tabCollectionManager,
                _tabItemsSelectionCoordinator,
                _tabNavigationController,
                _textEditorOverlayController,
                _toolWindowSessionManager!,
                () => _closedTabsRestorer?.IsRestoring == true,
                this.UpdateWindowTabsInfo
            );
            // Workspace synchronizer создаётся последним: он координирует уже готовые collection,
            // history, overlay, session и activation при событиях документов и solution.
            _tabsWorkspaceSynchronizer = new TMEx.Controls.Tabs.TabsWorkspaceSynchronizer(
                PackageServices.Dte2,
                this.Dispatcher,
                _tabCollectionManager,
                _closedTabsHistory,
                _textEditorOverlayController,
                _toolWindowSessionManager!,
                _tabActivationSynchronizer,
                this.UninitializeFileWatcher
            );
        }

        private void UninitializeTabItemsSelectionCoordinator() {
            _keyboardTabNavigationExtension.FocusedItemChanged -= this.OnKeyboardFocusedTabItemChanged;
            _keyboardTabNavigationExtension.InputTargetRestoreRequested -= this.OnKeyboardInputTargetRestoreRequested;
            _textEditorOverlayController.Dispose();
            _tabCommandController = null;
            _tabInputController = null;
            _tabsWorkspaceSynchronizer = null;
            _tabActivationSynchronizer = null;
            _virtualMenuController = null;
            _projectContextController = null;
            _closedTabsRestorer = null;
        }


        // 
        // ░ BackgroundRoutine
        //
        private void InitializeBackgroundRoutine() {
            _tabsStateReconciler = new TMEx.Controls.Tabs.TabsStateReconciler(
                PackageServices.Dte2,
                this.Dispatcher,
                _tabCollectionManager,
                this.UpdateWindowTabsInfo
            );
            _tabsStateReconciler.Start();
            _tabsWorkspaceSynchronizer?.SetStateReconciler(_tabsStateReconciler);
        }

        private void UninitializeBackgroundRoutine() {
            _tabsStateReconciler?.Dispose();
            _tabsStateReconciler = null;
        }


        //
        // ░ Event handlers
        // ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 
        // 
        // ░ DTE
        //
        private void OnDocumentOpened(EnvDTE.Document document) {
            _tabsWorkspaceSynchronizer?.HandleDocumentOpened(document);
        }


        private void OnDocumentSaved(EnvDTE.Document document) {
            _tabsWorkspaceSynchronizer?.HandleDocumentSaved(document);
        }


        private void OnDocumentClosing(EnvDTE.Document document) {
            _tabsWorkspaceSynchronizer?.HandleDocumentClosing(document);
        }
        

        private void OnWindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window lostFocus) {
            _tabActivationSynchronizer?.HandleWindowActivated(gotFocus);
        }


        private void OnWindowClosing(EnvDTE.Window closingWindow) {
            _tabsWorkspaceSynchronizer?.HandleWindowClosing(closingWindow);
        }


        private void OnSolutionClosing() {
            _tabsWorkspaceSynchronizer?.HandleSolutionClosing();
        }

        private void OnSolutionClosed(string solutionName) {
            _tabsWorkspaceSynchronizer?.HandleSolutionClosed(solutionName);
        }
        

        // 
        // ░ FileWatcher
        //
        private void OnFileChanged(string fullPath) {
            _tabsWorkspaceSynchronizer?.HandleFileChanged(fullPath);
        }

        private void OnFileRenamed(string oldFullPath, string newFullPath) {
            _tabsWorkspaceSynchronizer?.HandleFileRenamed(oldFullPath, newFullPath);
        }

        private void OnFileDeleted(string fullPath) {
            _tabsWorkspaceSynchronizer?.HandleFileDeleted(fullPath);
        }


        //
        // ░ VsShellTrackers 
        //
        private void OnDocumentActivatedExternally(VsShell._EventArgs.DocumentNavigationEventArgs e) {
            _tabActivationSynchronizer?.HandleDocumentActivatedExternally(e);
        }


        private void OnVsWindowFrameActivated(IVsWindowFrame vsWindowFrame) {
            _tabActivationSynchronizer?.HandleWindowFrameActivated(vsWindowFrame);
        }


        // 
        // ░ TabItemsSelectionCoordinator
        //
        private void OnTabItemSelectionChanged(TMEx.State.Document.TabItemsGroupBase group, TMEx.State.Document.TabItemBase tabItem, bool isSelected) {
            // Выделение здесь намеренно не активирует документ. Решение об активации принимает
            // TabNavigationController с учётом модификаторов и выбранной TabSelectionActivationPolicy.
            // Флаг помечает одно внешнее изменение от VS; после доставки события его нужно погасить,
            // чтобы следующий пользовательский выбор обрабатывался как обычный.
            if (tabItem.Metadata.GetFlag("IsActivatedExternally")) {
                tabItem.Metadata.SetFlag("IsActivatedExternally", false);
            }

            VsixThreadHelper.RunOnUiThread(
                this.Dispatcher,
                this.UpdateSelectionCount,
                DispatcherPriority.DataBind
            );
        }

        private void OnSelectionStateChanged(Helpers.Enums.SelectionState selectionState) {
            this.IsMultipleTabSelection = selectionState == Helpers.Enums.SelectionState.Multiple;
            this.UpdateSelectionCount();
        }

        private void UpdateSelectionCount() {
            int selectedCount = _tabItemsSelectionCoordinator.SelectedItems.Count;
            this.SelectionCountText.Text = $"{selectedCount} tabs";
            this.SelectionCountText.Visibility = selectedCount > 1 ? Visibility.Visible : Visibility.Collapsed;
        }


        // 
        // ░ BackgroundRoutine
        //
        private void TabsManagerStateTimerHandler(object sender, EventArgs e) {
            _tabsStateReconciler?.Reconcile();
        }


        //
        // ░ UI click handlers
        // ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 
        //
        // ░ Commands
        //
        private void OnPinTabItem(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _tabCommandController?.Pin(parameter);
        }


        private void OnUnpinTabItem(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _tabCommandController?.Unpin(parameter);
        }


        private void OnCloseTabItem(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _tabCommandController?.Close(parameter);
        }

        private void OnCloseSelectedTabItems(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _tabCommandController?.CloseSelected();
        }

        private void CloseTabItems(IReadOnlyList<TMEx.State.Document.TabItemBase> tabItems) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _tabCommandController?.Close(tabItems);
        }

        private void OnKeepOpenedTabItem(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _tabCommandController?.KeepOpen(parameter);
        }

        private void OnCopyTabName(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _tabCommandController?.CopyName(parameter);
        }

        private void OnCopyTabPath(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _tabCommandController?.CopyPath(parameter);
        }

        private void OnOpenLocationTabItem(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _tabCommandController?.OpenLocation(parameter);
        }

        private void OnMoveTabItemToRelatedProject(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _projectContextController?.MoveToRelatedProject(parameter);
        }

        private void OnMoveTabItemToRelatedProjectFile(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _projectContextController?.MoveToRelatedProjectFile(parameter);
        }

        private void OnProjectContextIncludersOpen(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _projectContextController?.ToggleIncludersMenu(parameter);
        }

        private void OnProjectContextMenuItemMouseEnter(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _projectContextController?.HandleMenuItemMouseEnter(parameter);
        }

        private void OnReloadDocumentReferencesProjects(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _projectContextController?.ReloadProjects(parameter);
        }


        //
        // ░ ContextMenu
        //
        private void OnTabItemContextMenuOpen(object parameter) {
            if (parameter is not Controls.MenuControl.MenuOpeningArgs contextMenuOpeningArgs ||
                contextMenuOpeningArgs.DataContext is not TMEx.State.Document.TabItemBase tabItem) {

                return;
            }

            if (_tabItemsSelectionCoordinator.SelectionState == Helpers.Enums.SelectionState.Single) {
                tabItem.Metadata.SetFlag("IsCtxMenuOpenned", true);
                this.ContextMenuItems = new ObservableCollection<Helpers.IMenuItem>(
                    _tabMenuItemFactory.CreateSingleSelectionContextMenu(tabItem)
                );
                return;
            }

            var selectedTabItems = _tabItemsSelectionCoordinator.SelectedItems
                .Select(entry => entry.Item)
                .ToList();
            bool isAnchorSelected = selectedTabItems.Any(selectedTabItem => ReferenceEquals(selectedTabItem, tabItem));
            if (isAnchorSelected) {
                this.ContextMenuItems = new ObservableCollection<Helpers.IMenuItem>(
                    _tabMenuItemFactory.CreateMultipleSelectionContextMenu(tabItem, selectedTabItems)
                );
            }
            else {
                contextMenuOpeningArgs.ShouldOpen = false;
                tabItem.IsSelected = true;
            }
        }

        private void OnTabItemContextMenuClosed(object parameter) {
            if (parameter is Controls.MenuControl.MenuClosedArgs contextMenuClosedArgs) {
                if (contextMenuClosedArgs.DataContext is TMEx.State.Document.TabItemBase tabItem) {
                    tabItem.Metadata.SetFlag("IsCtxMenuOpenned", false);
                }
            }
        }


        //
        // ░ VirtualMenu
        //
        private void CloseButton_MouseEnter(object sender, MouseEventArgs e) {
            _virtualMenuController?.HandleCloseButtonMouseEnter(sender);
        }

        private void InteractiveArea_MouseEnter(object sender, MouseEventArgs e) {
            _virtualMenuController?.HandleInteractiveAreaMouseEnter(sender);
        }

        private void InteractiveArea_MouseLeave(object sender, MouseEventArgs e) {
            _virtualMenuController?.HandleInteractiveAreaMouseLeave();
        }

        private void OnTabItemVirtualMenuOpen(object parameter) {
            _virtualMenuController?.Open(parameter);
        }

        private void OnTabItemVirtualMenuClosed(object parameter) {
            _virtualMenuController?.Close(parameter);
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
            double requestedHeight = this.ScaleFactorTabsCompactness * 26.4;
            // Новая база 100% равна прежней высоте при 120%: 22 * 1.2 = 26.4.
            Helpers.BaseUserControlResourceHelper.UpdateDynamicResource(this, "AppTabItemHeight", requestedHeight);
            _tabInputController?.ApplyScale();
        }


        //
        // ░ История закрытых вкладок
        //
        private TMEx.Controls.Tabs.ClosedTabEntry CreateClosedTabEntry(TMEx.State.Document.TabItemBase tabItem) {
            return _tabsWorkspaceSynchronizer!.CreateClosedTabEntry(tabItem);
        }

        private void RestoreLastClosedTabsOperation() {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!_closedTabsHistory.TryPop(out var operation) || operation == null) {
                return;
            }

            _closedTabsRestorer?.Restore(operation);
        }

        private void UpdateWindowTabsInfo() {
            _tabsWorkspaceSynchronizer?.UpdateWindowTabsInfo();
        }

        internal void RegisterTabItemControl(TabItemControl control) {
            _tabInputController?.Register(control);
        }

        internal void UnregisterTabItemControl(TabItemControl control) {
            _tabInputController?.Unregister(control);
        }

        internal bool HandleTabEditKey(Key key, ModifierKeys modifiers) {
            return _tabInputController?.HandleEditKey(key, modifiers) == true;
        }

        internal bool TryRenameTabItem(
            TabItemControl source,
            string proposedName,
            IReadOnlyList<TMEx.State.Document.TabItemDocument> renameGroupTabItems
            ) {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (source.DataContext is not TMEx.State.Document.TabItemDocument tabItemDocument) {
                this.ShowTabRenameMessage("Only document tabs can be renamed.");
                return false;
            }

            var result = _tabRenameService.Rename(tabItemDocument, proposedName, renameGroupTabItems);
            if (!result.Succeeded && result.ErrorMessage != null) {
                this.ShowTabRenameMessage(result.ErrorMessage);
            }

            return result.Succeeded;
        }

        internal void RestoreTabNavigationInputTarget() {
            if (this.IsTabEditMode) {
                _tabInputController?.RestoreInputTarget();
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
            if (key == Key.C) {
                return this.IsTabEditMode;
            }

            if (key == Key.Home || key == Key.End) {
                return Keyboard.FocusedElement is DependencyObject focusedElement &&
                    Helpers.VisualTree.FindParentByType<TabItemControl>(focusedElement)?.IsRenaming == true;
            }

            return _tabInputController?.CanHandleRedirectedKey(key) == true;
        }

        internal void HandleTabPointerNavigation(TabItemControl source, ModifierKeys modifiers) {
            _tabInputController?.HandlePointerNavigation(source, modifiers);
        }

        private List<TMEx.State.Document.TabItemBase> GetVisibleTabItems() {
            return _tabCollectionManager.GetSnapshot().ToList();
        }

        private void OnKeyboardFocusedTabItemChanged(TMEx.State.Document.TabItemBase? tabItem) {
            _tabInputController?.HandleFocusedItemChanged(tabItem);
        }

        private void OnKeyboardInputTargetRestoreRequested() {
            _tabInputController?.RestoreInputTarget();
        }

        private void FocusEditModeInputTarget() {
            _tabInputController?.FocusInputTarget();
        }

        private void ApplyAppearance() {
            _tabInputController?.ApplyAppearance();
        }

    }
}
