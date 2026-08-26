using System;
using System.Linq;
using System.Collections.Generic;

using TMEx = TabsManagerExtension;


namespace TabsManagerExtension.Controls.Tabs {
    internal enum ClosedTabKind {
        Document,
        ToolWindow
    }

    internal enum ClosedTabGroupKind {
        Default,
        Pinned,
        Preview
    }

    internal sealed class ClosedTabEntry {
        public ClosedTabKind Kind { get; set; }
        public string FullName { get; set; } = string.Empty;
        public string? WindowId { get; set; }
        public string GroupName { get; set; } = string.Empty;
        public ClosedTabGroupKind GroupKind { get; set; }
    }

    internal sealed class ClosedTabsOperation {
        public IReadOnlyList<ClosedTabEntry> Entries { get; }

        public ClosedTabsOperation(IReadOnlyList<ClosedTabEntry> entries) {
            this.Entries = entries;
        }
    }

    /// <summary>Хранит undo-операции закрытия и отличает командное закрытие от внешнего.</summary>
    internal sealed class ClosedTabsHistory {
        private const int Capacity = 50;

        private readonly Stack<ClosedTabsOperation> _operations = new();
        private readonly HashSet<string> _tabsBeingClosed = new(StringComparer.OrdinalIgnoreCase);

        public int Count => _operations.Count;

        public ClosedTabEntry CreateEntry(TMEx.State.Document.TabItemBase tabItem, TMEx.State.Document.TabItemsGroupBase? group) {
            return new ClosedTabEntry {
                Kind = tabItem is TMEx.State.Document.TabItemWindow ? ClosedTabKind.ToolWindow : ClosedTabKind.Document,
                FullName = tabItem.FullName,
                WindowId = (tabItem as TMEx.State.Document.TabItemWindow)?.WindowId,
                GroupName = group?.GroupName ?? string.Empty,
                GroupKind = GetGroupKind(group)
            };
        }

        public void MarkClosing(TMEx.State.Document.TabItemBase tabItem) {
            _tabsBeingClosed.Add(GetKey(tabItem));
        }

        public bool ConsumeClosingMark(TMEx.State.Document.TabItemBase tabItem) {
            return _tabsBeingClosed.Remove(GetKey(tabItem));
        }

        public void RemoveClosingMark(TMEx.State.Document.TabItemBase tabItem) {
            _tabsBeingClosed.Remove(GetKey(tabItem));
        }

        public void Push(IEnumerable<ClosedTabEntry> entries) {
            var operationEntries = entries.ToList();
            if (operationEntries.Count == 0) {
                return;
            }

            _operations.Push(new ClosedTabsOperation(operationEntries));
            if (_operations.Count <= Capacity) {
                return;
            }

            var retainedOperations = _operations.Take(Capacity).Reverse().ToList();
            _operations.Clear();
            foreach (var operation in retainedOperations) {
                _operations.Push(operation);
            }
        }

        public bool TryPop(out ClosedTabsOperation? operation) {
            if (_operations.Count == 0) {
                operation = null;
                return false;
            }

            operation = _operations.Pop();
            return true;
        }

        public void Clear() {
            _operations.Clear();
            _tabsBeingClosed.Clear();
        }

        private static ClosedTabGroupKind GetGroupKind(TMEx.State.Document.TabItemsGroupBase? group) {
            if (group is TMEx.State.Document.TabItemsPreviewGroup) {
                return ClosedTabGroupKind.Preview;
            }
            if (group is TMEx.State.Document.TabItemsPinnedGroup) {
                return ClosedTabGroupKind.Pinned;
            }

            return ClosedTabGroupKind.Default;
        }

        private static string GetKey(TMEx.State.Document.TabItemBase tabItem) {
            // Префикс исключает случайное совпадение пути документа с идентификатором окна.
            return tabItem is TMEx.State.Document.TabItemWindow tabItemWindow
                ? $"window:{tabItemWindow.WindowId}"
                : $"document:{tabItem.FullName}";
        }
    }
}
