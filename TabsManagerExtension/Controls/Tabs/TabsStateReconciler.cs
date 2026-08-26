using System;
using System.Linq;
using System.Windows.Threading;
using System.Collections.Generic;
using Microsoft.VisualStudio.Shell;

using TMEx = TabsManagerExtension;

namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Периодически сверяет модели вкладок с фактическим состоянием Visual Studio.</summary>
    internal sealed class TabsStateReconciler : IDisposable {
        private readonly EnvDTE80.DTE2 _dte;
        private readonly DispatcherTimer _timer;
        private readonly TabCollectionManager _tabCollectionManager;
        private readonly Action _onUpdateWindowTabsInfo;

        public TabsStateReconciler(
            EnvDTE80.DTE2 dte,
            Dispatcher dispatcher,
            TabCollectionManager tabCollectionManager,
            Action onUpdateWindowTabsInfo
            ) {
            _dte = dte;
            _tabCollectionManager = tabCollectionManager;
            _onUpdateWindowTabsInfo = onUpdateWindowTabsInfo;
            // DTE не гарантирует событие для каждого изменения preview/save/tool-window state.
            // Редкий UI-thread timer служит страховочной сверкой, а не основным event pipeline.
            _timer = new DispatcherTimer(DispatcherPriority.Normal, dispatcher) {
                Interval = TimeSpan.FromSeconds(2)
            };
            _timer.Tick += this.OnTimerTick;
        }

        public void Start() {
            _timer.Start();
        }

        public void Stop() {
            _timer.Stop();
        }

        public void Reconcile() {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Порядок важен: сначала актуализируем документы и удаляем исчезнувшие окна,
            // затем строим итоговый снимок tool windows для сохранения settings.
            this.UpdateDocumentSaveStates();
            this.MoveDocumentsOutOfPreviewGroup();
            this.RemoveClosedToolWindows();
            _onUpdateWindowTabsInfo();
        }

        public void Dispose() {
            _timer.Stop();
            _timer.Tick -= this.OnTimerTick;
        }

        private void OnTimerTick(object sender, EventArgs e) {
            this.Reconcile();
        }

        private void UpdateDocumentSaveStates() {
            // Снимок DTE-документов позволяет выполнить один регистронезависимый lookup на
            // вкладку и не перечислять COM-коллекцию заново для каждой группы.
            var openDocuments = _dte.Documents
                .Cast<EnvDTE.Document>()
                .ToDictionary(document => document.FullName, StringComparer.OrdinalIgnoreCase);

            foreach (var group in _tabCollectionManager.Groups.ToList()) {
                foreach (var tabItem in group.Items.ToList()) {
                    if (!openDocuments.TryGetValue(tabItem.FullName, out var document)) {
                        continue;
                    }

                    if (document.Saved) {
                        // Звёздочка является только UI-проекцией DTE.Document.Saved и потому
                        // восстанавливается сверкой, если соответствующее событие было пропущено.
                        tabItem.Caption = tabItem.Caption.TrimEnd('*');
                    }
                    else if (!tabItem.Caption.EndsWith("*")) {
                        tabItem.Caption += "*";
                    }
                }
            }
        }

        private void MoveDocumentsOutOfPreviewGroup() {
            var previewGroup = _tabCollectionManager.Groups.FirstOrDefault(group => group is TMEx.State.Document.TabItemsPreviewGroup);
            if (previewGroup == null) {
                return;
            }

            foreach (var document in previewGroup.Items.OfType<TMEx.State.Document.TabItemDocument>().ToList()) {
                // VS может закрепить preview-вкладку без отдельного события для нашей модели.
                // ShellDocument остаётся авторитетным источником фактического preview-state.
                if (!document.ShellDocument.IsDocumentInPreviewTab()) {
                    _tabCollectionManager.MovePreviewDocumentToDefaultGroup(document);
                }
            }
        }

        private void RemoveClosedToolWindows() {
            var openWindowIds = new HashSet<string>();
            try {
                // WindowId используется вместо Caption: заголовок окна может меняться, тогда
                // как persistence GUID остаётся стабильным между восстановлениями.
                openWindowIds = _dte.Windows
                    .Cast<EnvDTE.Window>()
                    .Select(VsShell.Document.ShellWindow.GetWindowId)
                    .Where(windowId => !string.IsNullOrEmpty(windowId))
                    .ToHashSet();
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to enumerate windows: {ex.Message}");
            }

            foreach (var group in _tabCollectionManager.Groups.ToList()) {
                var windowsToRemove = group.Items
                    .OfType<TMEx.State.Document.TabItemWindow>()
                    // Иногда DTE.Windows временно не содержит окно, которое ещё доступно через
                    // сохранённый COM-wrapper. Проверка Visible защищает от преждевременного удаления.
                    .Where(window => !openWindowIds.Contains(window.WindowId) && !IsWindowVisible(window))
                    .ToList();

                foreach (var window in windowsToRemove) {
                    _tabCollectionManager.Remove(window);
                }
            }
        }

        private static bool IsWindowVisible(TMEx.State.Document.TabItemWindow tabItemWindow) {
            ThreadHelper.ThrowIfNotOnUIThread();

            try {
                return tabItemWindow.ShellWindow.Window.Visible;
            }
            catch (Exception) {
                return false;
            }
        }
    }
}
