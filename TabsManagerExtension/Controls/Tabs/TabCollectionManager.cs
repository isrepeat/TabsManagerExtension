using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.VisualStudio.Shell;

using TMEx = TabsManagerExtension;

namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Единая точка поиска и изменения групп и моделей вкладок.</summary>
    internal sealed class TabCollectionManager {
        public Helpers.Collections.SortedObservableCollection<TMEx.State.Document.TabItemsGroupBase> Groups { get; }

        public IEnumerable<TMEx.State.Document.TabItemBase> AllTabs => this.Groups.SelectMany(group => group.Items);

        public TabCollectionManager() {
            // Обычные группы одного приоритета сортируются по отображаемому имени. Фактическое
            // расположение специальных групп задаётся ниже через PriorityGroup и не зависит от имени.
            var defaultGroupComparer = Comparer<TMEx.State.Document.TabItemsGroupBase>.Create(
                (left, right) => string.Compare(left.GroupName, right.GroupName, StringComparison.OrdinalIgnoreCase)
            );

            // Preview всегда один и заменяет предыдущую preview-группу. Разделители также
            // существуют в единственном экземпляре, а pinned/default могут содержать несколько
            // групп и сортируются внутри своего диапазона стандартным comparer.
            var priorityGroups = new List<Helpers.Collections.PriorityGroup<TMEx.State.Document.TabItemsGroupBase>> {
                new Helpers.Collections.PriorityGroup<TMEx.State.Document.TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top,
                    InsertMode = Helpers.Collections.ItemInsertMode.SingleWithReplaceExisting,
                    Predicate = group => group is TMEx.State.Document.TabItemsPreviewGroup,
                    Comparer = defaultGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TMEx.State.Document.TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top + 1,
                    InsertMode = Helpers.Collections.ItemInsertMode.Single,
                    Predicate = group => group is TMEx.State.Document.SeparatorTabItemsGroup separator && separator.Key == "Preview-Pinned",
                    Comparer = defaultGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TMEx.State.Document.TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top + 2,
                    Predicate = group => group is TMEx.State.Document.TabItemsPinnedGroup,
                    Comparer = defaultGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TMEx.State.Document.TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top + 3,
                    InsertMode = Helpers.Collections.ItemInsertMode.Single,
                    Predicate = group => group is TMEx.State.Document.SeparatorTabItemsGroup separator && separator.Key == "Pinned-Default",
                    Comparer = defaultGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TMEx.State.Document.TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Middle,
                    Predicate = group => group is TMEx.State.Document.TabItemsDefaultGroup,
                    Comparer = defaultGroupComparer
                }
            };

            this.Groups = new Helpers.Collections.SortedObservableCollection<TMEx.State.Document.TabItemsGroupBase>(
                defaultGroupComparer,
                priorityGroups
            );
        }

        public void AddDocumentToPreview(TMEx.State.Document.TabItemDocument tabItemDocument) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // SingleWithReplaceExisting у preview-группы обеспечивает семантику VS: одновременно
            // отображается только один preview-документ. AddToGroup возвращает канонический TabItem,
            // поэтому выделяем именно его, а не обязательно переданный экземпляр.
            var addedOrExistingTabItem = this.AddToGroup(tabItemDocument, new TMEx.State.Document.TabItemsPreviewGroup());
            addedOrExistingTabItem.IsSelected = true;
        }

        public TMEx.State.Document.TabItemBase AddAutomatically(TMEx.State.Document.TabItemBase tabItem) {
            ThreadHelper.ThrowIfNotOnUIThread();

            TMEx.State.Document.TabItemsGroupBase targetGroup = tabItem switch {
                TMEx.State.Document.TabItemDocument document => new TMEx.State.Document.TabItemsDefaultGroup(document.ShellDocument.GetDocumentProjectName()),
                TMEx.State.Document.TabItemWindow => new TMEx.State.Document.TabItemsDefaultGroup("[Tool Windows]"),
                _ => new TMEx.State.Document.TabItemsDefaultGroup("Other")
            };

            return this.AddToGroup(tabItem, targetGroup);
        }

        public TMEx.State.Document.TabItemBase AddToGroup(TMEx.State.Document.TabItemBase tabItem, TMEx.State.Document.TabItemsGroupBase targetGroup) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Один документ или tool window не должен одновременно находиться в нескольких
            // группах. При повторном DTE-событии сохраняем существующий объект вместе с selection
            // и metadata вместо добавления второй модели.
            var existingTabItem = this.Find(tabItem);
            if (existingTabItem != null) {
                return existingTabItem;
            }

            // Тип является частью идентичности группы: preview, pinned и default могут иметь
            // одинаковое отображаемое имя, но обладают разной семантикой.
            var existingGroup = this.Groups.FirstOrDefault(
                group => group.GetType() == targetGroup.GetType() && group.GroupName == targetGroup.GroupName
            );

            if (existingGroup == null) {
                this.Groups.Add(targetGroup);
                this.UpdateSeparators();
                existingGroup = targetGroup;
            }

            tabItem.IsPinnedTab = existingGroup is TMEx.State.Document.TabItemsPinnedGroup;
            if (tabItem is TMEx.State.Document.TabItemDocument tabItemDocument) {
                tabItemDocument.IsPreviewTab = existingGroup is TMEx.State.Document.TabItemsPreviewGroup;
            }

            Helpers.Diagnostic.Logger.LogDebug($"Added tab \"{tabItem.Caption}\" to group \"{existingGroup.GroupName}\": {tabItem}");
            existingGroup.Items.Add(tabItem);
            return tabItem;
        }

        public void MovePreviewDocumentToDefaultGroup(TMEx.State.Document.TabItemDocument tabItemDocument) {
            if (tabItemDocument == null || !tabItemDocument.IsPreviewTab) {
                return;
            }

            // Preview-группа по инварианту содержит единственную вкладку, поэтому удаляем всю
            // группу. Это заодно пересобирает разделители перед добавлением обычной вкладки.
            var previewGroup = this.Groups.FirstOrDefault(group => group is TMEx.State.Document.TabItemsPreviewGroup);
            if (previewGroup != null) {
                this.RemoveGroup(previewGroup);
            }

            this.AddAutomatically(tabItemDocument);
        }

        public bool Move(TMEx.State.Document.TabItemBase tabItem, TMEx.State.Document.TabItemsGroupBase targetGroup) {
            var current = this.FindWithGroup(tabItem);
            if (current == null) {
                return false;
            }

            // Переносим найденный канонический экземпляр, а не объект аргумента: вызывающий код
            // мог передать эквивалентную временную модель, найденную по пути или WindowId.
            this.RemoveFromGroup(current.Value.Item, current.Value.Group);
            this.AddToGroup(current.Value.Item, targetGroup);
            return true;
        }

        public void Remove(TMEx.State.Document.TabItemBase tabItem) {
            foreach (var group in this.Groups.ToList()) {
                if (group.Items.Contains(tabItem)) {
                    this.RemoveFromGroup(tabItem, group);
                    return;
                }
            }
        }

        public void RemoveFromGroup(TMEx.State.Document.TabItemBase tabItem, TMEx.State.Document.TabItemsGroupBase group) {
            if (!group.Items.Remove(tabItem)) {
                return;
            }

            Helpers.Diagnostic.Logger.LogDebug($"Removed tab \"{tabItem.Caption}\" from group \"{group.GroupName}\"");
            if (!group.Items.Any()) {
                this.RemoveGroup(group);
            }
        }

        public void RemoveGroup(TMEx.State.Document.TabItemsGroupBase group) {
            if (!this.Groups.Remove(group)) {
                return;
            }

            Helpers.Diagnostic.Logger.LogDebug($"Removed group \"{group.GroupName}\"");
            this.UpdateSeparators();
        }

        public bool HasGroup<T>() where T : TMEx.State.Document.TabItemsGroupBase {
            return this.Groups.OfType<T>().Any();
        }

        public IReadOnlyList<TMEx.State.Document.TabItemBase> GetSnapshot() {
            return this.AllTabs.ToList();
        }

        public void Clear() {
            this.Groups.Clear();
        }

        public void SetActiveFrame(TMEx.State.Document.TabItemBase activeTabItem) {
            // IsFrameActive — взаимоисключающий маркер реального активного VS frame. Он хранится
            // отдельно от IsSelected, поскольку при мультивыборе выделено несколько вкладок.
            foreach (var tabItem in this.AllTabs) {
                tabItem.Metadata?.SetFlag("IsFrameActive", ReferenceEquals(tabItem, activeTabItem));
            }
        }

        public void UpdateDocumentPath(string oldPath, string? newPath = null) {
            var tabItem = this.AllTabs.FirstOrDefault(
                item => string.Equals(item.FullName, oldPath, StringComparison.OrdinalIgnoreCase)
            );
            if (tabItem == null) {
                return;
            }

            if (newPath == null) {
                // FileSystemWatcher использует вызов без newPath как invalidation: путь остаётся
                // прежним, но caption перечитывается после внешнего изменения файла.
                tabItem.Caption = System.IO.Path.GetFileName(oldPath);
                return;
            }

            tabItem.FullName = newPath;
            tabItem.Caption = System.IO.Path.GetFileName(newPath);
        }

        public TMEx.State.Document.TabItemDocument? Find(EnvDTE.Document document) {
            ThreadHelper.ThrowIfNotOnUIThread();
            return this.Find(document.FullName);
        }

        public TMEx.State.Document.TabItemWindow? Find(EnvDTE.Window window) {
            // Временный TabItemWindow нужен только для вычисления стабильного WindowId.
            // Сравнение ссылок на DTE.Window ненадёжно после восстановления tool window.
            var result = this.FindWithGroup(new TMEx.State.Document.TabItemWindow(window));
            return result?.Item as TMEx.State.Document.TabItemWindow;
        }

        public TMEx.State.Document.TabItemBase? Find(TMEx.State.Document.TabItemBase tabItem) {
            return this.FindWithGroup(tabItem)?.Item;
        }

        public TMEx.State.Document.TabItemDocument? Find(string documentFullName) {
            return this.FindWithGroup(documentFullName)?.Item;
        }

        public (TMEx.State.Document.TabItemDocument Item, TMEx.State.Document.TabItemsGroupBase Group)? FindWithGroup(EnvDTE.Document document) {
            ThreadHelper.ThrowIfNotOnUIThread();
            return this.FindWithGroup(document.FullName);
        }

        public (TMEx.State.Document.TabItemWindow Item, TMEx.State.Document.TabItemsGroupBase Group)? FindWithGroup(EnvDTE.Window window) {
            var result = this.FindWithGroup(new TMEx.State.Document.TabItemWindow(window));
            if (result is { Item: TMEx.State.Document.TabItemWindow tabItemWindow, Group: var group }) {
                return (tabItemWindow, group);
            }

            return null;
        }

        public (TMEx.State.Document.TabItemBase Item, TMEx.State.Document.TabItemsGroupBase Group)? FindWithGroup(TMEx.State.Document.TabItemBase tabItem) {
            if (tabItem is TMEx.State.Document.TabItemWindow tabItemWindow) {
                return this.FindWithGroupBy<TMEx.State.Document.TabItemWindow>(
                    window => string.Equals(window.WindowId, tabItemWindow.WindowId, StringComparison.OrdinalIgnoreCase)
                );
            }

            return this.FindWithGroupBy<TMEx.State.Document.TabItemBase>(
                item => string.Equals(item.FullName, tabItem.FullName, StringComparison.OrdinalIgnoreCase)
            );
        }

        public (TMEx.State.Document.TabItemDocument Item, TMEx.State.Document.TabItemsGroupBase Group)? FindWithGroup(string documentFullName) {
            return this.FindWithGroupBy<TMEx.State.Document.TabItemDocument>(
                document => string.Equals(document.FullName, documentFullName, StringComparison.OrdinalIgnoreCase)
            );
        }

        public void ForEach<T>(Action<T> action) where T : TMEx.State.Document.TabItemBase {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (var group in this.Groups) {
                foreach (var tabItem in group.Items.OfType<T>()) {
                    action(tabItem);
                }
            }
        }

        private (T Item, TMEx.State.Document.TabItemsGroupBase Group)? FindWithGroupBy<T>(Func<T, bool> predicate) where T : TMEx.State.Document.TabItemBase {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (var group in this.Groups) {
                var match = group.Items.OfType<T>().FirstOrDefault(predicate);
                if (match != null) {
                    return (match, group);
                }
            }

            return null;
        }

        private void UpdateSeparators() {
            // Проще и надёжнее пересобрать два производных separator-элемента после изменения
            // состава групп, чем поддерживать их инкрементально для каждой комбинации операций.
            foreach (var separator in this.Groups.OfType<TMEx.State.Document.SeparatorTabItemsGroup>().ToList()) {
                this.Groups.Remove(separator);
            }

            // Разделитель нужен только между реально существующими соседними секциями: пустая
            // preview/pinned/default-секция не должна оставлять декоративную строку в UI.
            if (this.HasGroup<TMEx.State.Document.TabItemsPreviewGroup>() &&
                (this.HasGroup<TMEx.State.Document.TabItemsPinnedGroup>() || this.HasGroup<TMEx.State.Document.TabItemsDefaultGroup>())) {

                this.Groups.Add(new TMEx.State.Document.SeparatorTabItemsGroup("Preview-Pinned"));
            }

            if (this.HasGroup<TMEx.State.Document.TabItemsPinnedGroup>() && this.HasGroup<TMEx.State.Document.TabItemsDefaultGroup>()) {
                this.Groups.Add(new TMEx.State.Document.SeparatorTabItemsGroup("Pinned-Default"));
            }
        }
    }
}
