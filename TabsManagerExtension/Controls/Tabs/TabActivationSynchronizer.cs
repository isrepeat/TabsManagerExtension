using System;
using System.Windows.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;

using TMEx = TabsManagerExtension;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>
    /// Сводит источники активации VS к единому active frame, selection и editor overlay.
    /// </summary>
    internal sealed class TabActivationSynchronizer {
        private readonly EnvDTE80.DTE2 _dte;
        private readonly Dispatcher _dispatcher;
        private readonly TabCollectionManager _tabCollectionManager;
        private readonly Helpers.Collections.GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase> _selectionCoordinator;
        private readonly Navigation.TabNavigationController _navigationController;
        private readonly VsShell.TextEditor.Overlay.TextEditorOverlayController _overlayController;
        private readonly ToolWindowSessionManager _toolWindowSessionManager;
        private readonly Func<bool> _isRestoringClosedTabs;
        private readonly Action _onUpdateWindowTabsInfo;

        public TabActivationSynchronizer(
            EnvDTE80.DTE2 dte,
            Dispatcher dispatcher,
            TabCollectionManager tabCollectionManager,
            Helpers.Collections.GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase> selectionCoordinator,
            Navigation.TabNavigationController navigationController,
            VsShell.TextEditor.Overlay.TextEditorOverlayController overlayController,
            ToolWindowSessionManager toolWindowSessionManager,
            Func<bool> isRestoringClosedTabs,
            Action onUpdateWindowTabsInfo
            ) {
            _dte = dte;
            _dispatcher = dispatcher;
            _tabCollectionManager = tabCollectionManager;
            _selectionCoordinator = selectionCoordinator;
            _navigationController = navigationController;
            _overlayController = overlayController;
            _toolWindowSessionManager = toolWindowSessionManager;
            _isRestoringClosedTabs = isRestoringClosedTabs;
            _onUpdateWindowTabsInfo = onUpdateWindowTabsInfo;
        }

        public void HandleWindowActivated(EnvDTE.Window gotFocus) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (gotFocus == null) {
                return;
            }

            var shellWindow = new VsShell.Document.ShellWindow(gotFocus);
            if (!shellWindow.IsTabWindow()) {
                return;
            }

            // Во время восстановления tool window Visual Studio может кратковременно вернуть
            // фокус документу. Не принимаем этот промежуточный фокус за выбор пользователя.
            if (gotFocus.Document != null && _toolWindowSessionManager.KeepPendingWindowActive()) {
                return;
            }

            // DTE одинаково сообщает об активации документов и tool window. Документ может
            // появиться раньше события открытия, поэтому при необходимости создаём его модель.
            TMEx.State.Document.TabItemBase tabItem;
            if (gotFocus.Document != null) {
                tabItem = _tabCollectionManager.Find(gotFocus.Document.FullName) ??
                    _tabCollectionManager.AddAutomatically(new TMEx.State.Document.TabItemDocument(gotFocus.Document));
            }
            else {
                tabItem = _tabCollectionManager.Find(gotFocus) ??
                    _tabCollectionManager.AddAutomatically(new TMEx.State.Document.TabItemWindow(shellWindow));
            }

            _tabCollectionManager.SetActiveFrame(tabItem);
            this.SelectActivated(tabItem);
            if (tabItem is TMEx.State.Document.TabItemWindow) {
                // Состав и активность tool window сохраняются отдельно от документов решения.
                _onUpdateWindowTabsInfo();
                _toolWindowSessionManager.Save(VsShell.Solution.Services.VsSolutionEventsTrackerService.Instance.IsSolutionClosing);
            }

            if (tabItem is TMEx.State.Document.TabItemDocument) {
                // Overlay показываем после завершения текущей обработки активации: к этому
                // моменту Visual Studio успевает установить актуальный editor frame.
                VsixThreadHelper.RunOnUiThread(
                    _dispatcher,
                    _overlayController.Show,
                    DispatcherPriority.Background
                );
            }
            else {
                Helpers.GlobalFlags.SetFlag("TextEditorFrameFocused", false);
                _overlayController.DeactivateEditorFrame();
            }
        }

        public void HandleDocumentActivatedExternally(VsShell._EventArgs.DocumentNavigationEventArgs eventArgs) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var tabItem = _tabCollectionManager.Find(eventArgs.CurrentDocumentFullName);
            if (tabItem == null) {
                return;
            }

            _tabCollectionManager.SetActiveFrame(tabItem);
            // Флаг позволяет selection coordinator отличить синхронизацию от VS
            // от пользовательского клика по вкладке и не активировать документ повторно.
            tabItem.Metadata?.SetFlag("IsActivatedExternally", true);
            this.SelectActivated(tabItem);
        }

        public void HandleWindowFrameActivated(IVsWindowFrame windowFrame) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // IVsWindowFrame даёт более точную информацию о типе активного frame, чем DTE:
            // MDI child соответствует документу редактора, остальные режимы — tool window.
            windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_FrameMode, out var mode);
            bool isMdiChild = mode != null && (VSFRAMEMODE)(int)mode == VSFRAMEMODE.VSFM_MdiChild;
            Helpers.GlobalFlags.SetFlag("TextEditorFrameFocused", isMdiChild);

            windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_ExtWindowObject, out var activatedWindowObject);
            var activeWindow = activatedWindowObject as EnvDTE.Window;
            // Сохраняем тот же приоритет незавершённого восстановления, что и в DTE-событии.
            if (activeWindow?.Document != null && _toolWindowSessionManager.KeepPendingWindowActive()) {
                return;
            }

            // Не создаём модель здесь: frame-событие может относиться к служебному окну VS.
            // Добавление настоящих вкладок выполняется обработчиками открытия и DTE-активации.
            TMEx.State.Document.TabItemBase? activeFrameTabItem = activeWindow?.Document != null
                ? _tabCollectionManager.Find(activeWindow.Document)
                : activeWindow == null ? null : _tabCollectionManager.Find(activeWindow);
            if (activeFrameTabItem != null) {
                _tabCollectionManager.SetActiveFrame(activeFrameTabItem);
                this.SelectActivated(activeFrameTabItem);
            }

            if (isMdiChild) {
                // Overlay привязывается к конкретному IVsTextView, а не только к DTE.Document:
                // это важно при нескольких представлениях одного документа.
                windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out var documentView);
                if (documentView is IVsCodeWindow codeWindow &&
                    codeWindow.GetPrimaryView(out var textView) == VSConstants.S_OK &&
                    textView != null) {

                    _overlayController.ActivateEditorFrame(textView);
                    return;
                }
            }

            _overlayController.DeactivateEditorFrame();
        }

        public void ActivatePrimaryTab() {
            ThreadHelper.ThrowIfNotOnUIThread();

            var primaryTabItem = _selectionCoordinator.PrimarySelection?.Item;
            if (primaryTabItem is TMEx.State.Document.IActivatableTab activatableTab) {
                Helpers.Diagnostic.Logger.LogDebug($"Activate - \"{primaryTabItem.Caption}\"");
                activatableTab.Activate();
            }
        }

        public void SyncWithActiveWindow() {
            ThreadHelper.ThrowIfNotOnUIThread();

            var activeWindow = _dte.ActiveWindow;
            if (activeWindow == null) {
                return;
            }

            // Сначала обновляем только маркер фактически активного VS frame. Он не всегда
            // совпадает с PrimarySelection, например при сохранённом мультивыборе вкладок.
            TMEx.State.Document.TabItemBase? activeFrameTabItem = null;
            if (VsShell.Document.ShellWindow.IsTabWindow(activeWindow)) {
                activeFrameTabItem = activeWindow.Document != null
                    ? _tabCollectionManager.Find(activeWindow.Document)
                    : _tabCollectionManager.Find(activeWindow);
            }
            else if (_dte.ActiveDocument != null) {
                activeFrameTabItem = _tabCollectionManager.Find(_dte.ActiveDocument);
            }

            if (activeFrameTabItem != null) {
                _tabCollectionManager.SetActiveFrame(activeFrameTabItem);
            }

            var selectedTabItem = _selectionCoordinator.PrimarySelection?.Item;
            TMEx.State.Document.TabItemBase? targetTabItem;
            // Затем при необходимости подтягиваем selection к активному окну. Сравнение путей
            // и caption до поиска предотвращает лишний цикл активации и повторные события VS.
            if (VsShell.Document.ShellWindow.IsTabWindow(activeWindow)) {
                if (activeWindow.Document == null) {
                    if (string.Equals(activeWindow.Caption, selectedTabItem?.Caption, StringComparison.OrdinalIgnoreCase)) {
                        return;
                    }

                    targetTabItem = _tabCollectionManager.Find(activeWindow);
                }
                else {
                    if (string.Equals(activeWindow.Document.FullName, selectedTabItem?.FullName, StringComparison.OrdinalIgnoreCase)) {
                        return;
                    }

                    targetTabItem = _tabCollectionManager.Find(activeWindow.Document);
                }
            }
            else {
                var activeDocument = _dte.ActiveDocument;
                if (activeDocument == null ||
                    string.Equals(activeDocument.FullName, selectedTabItem?.FullName, StringComparison.OrdinalIgnoreCase)) {

                    return;
                }

                targetTabItem = _tabCollectionManager.Find(activeDocument);
            }

            if (targetTabItem != null) {
                this.SelectActivated(targetTabItem);
            }
        }

        public void SelectActivated(TMEx.State.Document.TabItemBase tabItem) {
            if (_isRestoringClosedTabs()) {
                // При пакетном восстановлении обычный IsSelected может активировать каждый
                // документ по очереди. Обновляем selection без повторной активации в VS.
                _navigationController.SetSelectionWithoutActivation(tabItem, true, System.Windows.Input.ModifierKeys.None);
                return;
            }

            // В обычном сценарии selection coordinator сам применит правила одиночного
            // или множественного выбора и обновит визуальное состояние групп.
            tabItem.IsSelected = true;
        }
    }
}
