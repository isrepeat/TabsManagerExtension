using System;
using System.Linq;
using System.Collections.Generic;

using Microsoft.VisualStudio.Shell;

using TabsManagerExtension.State.Document;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Единая точка поиска и изменения групп и моделей вкладок.</summary>
    internal sealed class TabCollectionManager {
        public Helpers.Collections.SortedObservableCollection<TabItemsGroupBase> Groups { get; }

        public IEnumerable<TabItemBase> AllTabs => this.Groups.SelectMany(group => group.Items);

        public TabCollectionManager() {
            var defaultGroupComparer = Comparer<TabItemsGroupBase>.Create(
                (left, right) => string.Compare(left.GroupName, right.GroupName, StringComparison.OrdinalIgnoreCase)
            );

            var priorityGroups = new List<Helpers.Collections.PriorityGroup<TabItemsGroupBase>> {
                new Helpers.Collections.PriorityGroup<TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top,
                    InsertMode = Helpers.Collections.ItemInsertMode.SingleWithReplaceExisting,
                    Predicate = group => group is TabItemsPreviewGroup,
                    Comparer = defaultGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top + 1,
                    InsertMode = Helpers.Collections.ItemInsertMode.Single,
                    Predicate = group => group is SeparatorTabItemsGroup separator && separator.Key == "Preview-Pinned",
                    Comparer = defaultGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top + 2,
                    Predicate = group => group is TabItemsPinnedGroup,
                    Comparer = defaultGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Top + 3,
                    InsertMode = Helpers.Collections.ItemInsertMode.Single,
                    Predicate = group => group is SeparatorTabItemsGroup separator && separator.Key == "Pinned-Default",
                    Comparer = defaultGroupComparer
                },
                new Helpers.Collections.PriorityGroup<TabItemsGroupBase> {
                    Position = Helpers.Collections.ItemPosition.Middle,
                    Predicate = group => group is TabItemsDefaultGroup,
                    Comparer = defaultGroupComparer
                }
            };

            this.Groups = new Helpers.Collections.SortedObservableCollection<TabItemsGroupBase>(
                defaultGroupComparer,
                priorityGroups
            );
        }

        public void AddDocumentToPreview(TabItemDocument tabItemDocument) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var addedOrExistingTabItem = this.AddToGroup(tabItemDocument, new TabItemsPreviewGroup());
            addedOrExistingTabItem.IsSelected = true;
        }

        public TabItemBase AddAutomatically(TabItemBase tabItem) {
            ThreadHelper.ThrowIfNotOnUIThread();

            TabItemsGroupBase targetGroup = tabItem switch {
                TabItemDocument document => new TabItemsDefaultGroup(document.ShellDocument.GetDocumentProjectName()),
                TabItemWindow => new TabItemsDefaultGroup("[Tool Windows]"),
                _ => new TabItemsDefaultGroup("Other")
            };

            return this.AddToGroup(tabItem, targetGroup);
        }

        public TabItemBase AddToGroup(TabItemBase tabItem, TabItemsGroupBase targetGroup) {
            ThreadHelper.ThrowIfNotOnUIThread();

            var existingTabItem = this.Find(tabItem);
            if (existingTabItem != null) {
                return existingTabItem;
            }

            var existingGroup = this.Groups.FirstOrDefault(
                group => group.GetType() == targetGroup.GetType() && group.GroupName == targetGroup.GroupName
            );

            if (existingGroup == null) {
                this.Groups.Add(targetGroup);
                this.UpdateSeparators();
                existingGroup = targetGroup;
            }

            tabItem.IsPinnedTab = existingGroup is TabItemsPinnedGroup;
            if (tabItem is TabItemDocument tabItemDocument) {
                tabItemDocument.IsPreviewTab = existingGroup is TabItemsPreviewGroup;
            }

            Helpers.Diagnostic.Logger.LogDebug($"Added tab \"{tabItem.Caption}\" to group \"{existingGroup.GroupName}\": {tabItem}");
            existingGroup.Items.Add(tabItem);
            return tabItem;
        }

        public void MovePreviewDocumentToDefaultGroup(TabItemDocument tabItemDocument) {
            if (tabItemDocument == null || !tabItemDocument.IsPreviewTab) {
                return;
            }

            var previewGroup = this.Groups.FirstOrDefault(group => group is TabItemsPreviewGroup);
            if (previewGroup != null) {
                this.RemoveGroup(previewGroup);
            }

            this.AddAutomatically(tabItemDocument);
        }

        public bool Move(TabItemBase tabItem, TabItemsGroupBase targetGroup) {
            var current = this.FindWithGroup(tabItem);
            if (current == null) {
                return false;
            }

            this.RemoveFromGroup(current.Value.Item, current.Value.Group);
            this.AddToGroup(current.Value.Item, targetGroup);
            return true;
        }

        public void Remove(TabItemBase tabItem) {
            foreach (var group in this.Groups.ToList()) {
                if (group.Items.Contains(tabItem)) {
                    this.RemoveFromGroup(tabItem, group);
                    return;
                }
            }
        }

        public void RemoveFromGroup(TabItemBase tabItem, TabItemsGroupBase group) {
            if (!group.Items.Remove(tabItem)) {
                return;
            }

            Helpers.Diagnostic.Logger.LogDebug($"Removed tab \"{tabItem.Caption}\" from group \"{group.GroupName}\"");
            if (!group.Items.Any()) {
                this.RemoveGroup(group);
            }
        }

        public void RemoveGroup(TabItemsGroupBase group) {
            if (!this.Groups.Remove(group)) {
                return;
            }

            Helpers.Diagnostic.Logger.LogDebug($"Removed group \"{group.GroupName}\"");
            this.UpdateSeparators();
        }

        public bool HasGroup<T>() where T : TabItemsGroupBase {
            return this.Groups.OfType<T>().Any();
        }

        public IReadOnlyList<TabItemBase> GetSnapshot() {
            return this.AllTabs.ToList();
        }

        public void Clear() {
            this.Groups.Clear();
        }

        public void SetActiveFrame(TabItemBase activeTabItem) {
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
                tabItem.Caption = System.IO.Path.GetFileName(oldPath);
                return;
            }

            tabItem.FullName = newPath;
            tabItem.Caption = System.IO.Path.GetFileName(newPath);
        }

        public TabItemDocument? Find(EnvDTE.Document document) {
            ThreadHelper.ThrowIfNotOnUIThread();
            return this.Find(document.FullName);
        }

        public TabItemWindow? Find(EnvDTE.Window window) {
            var result = this.FindWithGroup(new TabItemWindow(window));
            return result?.Item as TabItemWindow;
        }

        public TabItemBase? Find(TabItemBase tabItem) {
            return this.FindWithGroup(tabItem)?.Item;
        }

        public TabItemDocument? Find(string documentFullName) {
            return this.FindWithGroup(documentFullName)?.Item;
        }

        public (TabItemDocument Item, TabItemsGroupBase Group)? FindWithGroup(EnvDTE.Document document) {
            ThreadHelper.ThrowIfNotOnUIThread();
            return this.FindWithGroup(document.FullName);
        }

        public (TabItemWindow Item, TabItemsGroupBase Group)? FindWithGroup(EnvDTE.Window window) {
            var result = this.FindWithGroup(new TabItemWindow(window));
            if (result is { Item: TabItemWindow tabItemWindow, Group: var group }) {
                return (tabItemWindow, group);
            }

            return null;
        }

        public (TabItemBase Item, TabItemsGroupBase Group)? FindWithGroup(TabItemBase tabItem) {
            if (tabItem is TabItemWindow tabItemWindow) {
                return this.FindWithGroupBy<TabItemWindow>(
                    window => string.Equals(window.WindowId, tabItemWindow.WindowId, StringComparison.OrdinalIgnoreCase)
                );
            }

            return this.FindWithGroupBy<TabItemBase>(
                item => string.Equals(item.FullName, tabItem.FullName, StringComparison.OrdinalIgnoreCase)
            );
        }

        public (TabItemDocument Item, TabItemsGroupBase Group)? FindWithGroup(string documentFullName) {
            return this.FindWithGroupBy<TabItemDocument>(
                document => string.Equals(document.FullName, documentFullName, StringComparison.OrdinalIgnoreCase)
            );
        }

        public void ForEach<T>(Action<T> action) where T : TabItemBase {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (var group in this.Groups) {
                foreach (var tabItem in group.Items.OfType<T>()) {
                    action(tabItem);
                }
            }
        }

        private (T Item, TabItemsGroupBase Group)? FindWithGroupBy<T>(Func<T, bool> predicate) where T : TabItemBase {
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
            foreach (var separator in this.Groups.OfType<SeparatorTabItemsGroup>().ToList()) {
                this.Groups.Remove(separator);
            }

            if (this.HasGroup<TabItemsPreviewGroup>() &&
                (this.HasGroup<TabItemsPinnedGroup>() || this.HasGroup<TabItemsDefaultGroup>())) {

                this.Groups.Add(new SeparatorTabItemsGroup("Preview-Pinned"));
            }

            if (this.HasGroup<TabItemsPinnedGroup>() && this.HasGroup<TabItemsDefaultGroup>()) {
                this.Groups.Add(new SeparatorTabItemsGroup("Pinned-Default"));
            }
        }
    }
}
