using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using TMEx = TabsManagerExtension;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Открывает закрытые вкладки и возвращает их в сохранённые группы.</summary>
    internal sealed class ClosedTabsRestorer {
        private readonly EnvDTE80.DTE2 _dte;
        private readonly TabCollectionManager _tabCollectionManager;
        private readonly Action _onUpdateWindowTabsInfo;
        private readonly Action _onRestoreInputTarget;
        private readonly Action _onFocusInputTarget;

        public bool IsRestoring { get; private set; }

        public ClosedTabsRestorer(
            EnvDTE80.DTE2 dte,
            TabCollectionManager tabCollectionManager,
            Action onUpdateWindowTabsInfo,
            Action onRestoreInputTarget,
            Action onFocusInputTarget
            ) {
            _dte = dte;
            _tabCollectionManager = tabCollectionManager;
            _onUpdateWindowTabsInfo = onUpdateWindowTabsInfo;
            _onRestoreInputTarget = onRestoreInputTarget;
            _onFocusInputTarget = onFocusInputTarget;
        }

        public void Restore(ClosedTabsOperation operation) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Пока пакет восстанавливается, события DTE всё равно проходят через обычную цепочку
            // активации. Флаг заставляет TabActivationSynchronizer менять selection без повторной
            // активации каждого по очереди открываемого документа.
            this.IsRestoring = true;
            try {
                // Одна undo-операция может содержать мультивыбор. Ошибка отдельной вкладки
                // не должна отменять восстановление остальных элементов пакета.
                foreach (var entry in operation.Entries) {
                    try {
                        TMEx.State.Document.TabItemBase? restoredTabItem = entry.Kind == ClosedTabKind.Document
                            ? this.RestoreDocument(entry)
                            : this.RestoreToolWindow(entry);

                        if (restoredTabItem != null) {
                            this.MoveToOriginalGroup(restoredTabItem, entry);
                        }
                    }
                    catch (Exception ex) {
                        Helpers.Diagnostic.Logger.LogError($"Failed to restore closed tab '{entry.FullName}': {ex}");
                    }
                }
            }
            finally {
                this.IsRestoring = false;
            }

            // OpenFile/Show передают command target редактору или tool window. После завершения
            // всего пакета возвращаем клавиатурную навигацию панели только один раз.
            _onFocusInputTarget();
            _onRestoreInputTarget();
        }

        private TMEx.State.Document.TabItemDocument? RestoreDocument(ClosedTabEntry entry) {
            // Документ мог быть повторно открыт другим действием между Close и Ctrl+Z.
            // В этом случае достаточно вернуть уже существующую модель в исходную группу.
            var existingTabItem = _tabCollectionManager.Find(entry.FullName);
            if (existingTabItem != null) {
                return existingTabItem;
            }
            if (!File.Exists(entry.FullName)) {
                Helpers.Diagnostic.Logger.LogWarning($"Cannot restore deleted document '{entry.FullName}'");
                return null;
            }

            var window = _dte.ItemOperations.OpenFile(entry.FullName);
            // DocumentOpened обычно синхронно добавляет TabItem, но конкретный editor factory
            // может вернуть новый DTE.Document раньше обработчика события. Сначала ищем модель
            // по объекту возвращённого окна, затем используем стабильный moniker как fallback.
            var restoredTabItem = window?.Document == null ? null : _tabCollectionManager.Find(window.Document);
            return restoredTabItem ?? _tabCollectionManager.Find(entry.FullName);
        }

        private TMEx.State.Document.TabItemWindow? RestoreToolWindow(ClosedTabEntry entry) {
            // В отличие от документов, tool window восстанавливается по persistence GUID,
            // сохранённому Visual Studio для конкретного типа окна.
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
            // Show инициирует VS/DTE-события, которые создают либо обновляют TabItemWindow.
            // Объект EnvDTE.Window после восстановления может быть новым, поэтому обновляем
            // captions и находим модель по устойчивому WindowId, а не по ссылке на COM-объект.
            _onUpdateWindowTabsInfo();
            return _tabCollectionManager.Groups
                .SelectMany(group => group.Items)
                .OfType<TMEx.State.Document.TabItemWindow>()
                .FirstOrDefault(item => string.Equals(item.WindowId, entry.WindowId, StringComparison.OrdinalIgnoreCase));
        }

        private void MoveToOriginalGroup(TMEx.State.Document.TabItemBase tabItem, ClosedTabEntry entry) {
            // Обработчик открытия уже успел автоматически классифицировать вкладку. Сначала
            // удаляем эту временную классификацию, чтобы одна модель не осталась в двух группах.
            var current = _tabCollectionManager.FindWithGroup(tabItem);
            if (current != null) {
                _tabCollectionManager.RemoveFromGroup(current.Value.Item, current.Value.Group);
            }

            _tabCollectionManager.AddToGroup(tabItem, CreateGroup(entry));
            // Группа Tabs Manager хранит визуальную классификацию, а pinned-состояние документа
            // принадлежит ещё и native frame Visual Studio — его нужно восстановить отдельно.
            if (tabItem is TMEx.State.Document.TabItemDocument tabItemDocument && entry.GroupKind == ClosedTabGroupKind.Pinned) {
                tabItemDocument.ShellDocument.OpenDocumentAsPinned();
            }
        }

        private static TMEx.State.Document.TabItemsGroupBase CreateGroup(ClosedTabEntry entry) {
            // Preview-группа глобальная и не использует имя; default/pinned-группы восстанавливают
            // сохранённый project/group name из снимка, созданного до синхронного DocumentClosing.
            return entry.GroupKind switch {
                ClosedTabGroupKind.Preview => new TMEx.State.Document.TabItemsPreviewGroup(),
                ClosedTabGroupKind.Pinned => new TMEx.State.Document.TabItemsPinnedGroup(entry.GroupName),
                _ => new TMEx.State.Document.TabItemsDefaultGroup(entry.GroupName)
            };
        }
    }
}
