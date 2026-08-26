using System;
using System.Collections.Generic;

using TMEx = TabsManagerExtension;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Формирует пункты context и virtual menu из команд вкладок.</summary>
    internal sealed class TabMenuItemFactory {
        private readonly Helpers.RelayCommand<object> _pinCommand;
        private readonly Helpers.RelayCommand<object> _copyNameCommand;
        private readonly Helpers.RelayCommand<object> _copyPathCommand;
        private readonly Helpers.RelayCommand<object> _openLocationCommand;
        private readonly Helpers.RelayCommand<object> _closeCommand;
        private readonly Helpers.RelayCommand<object> _closeSelectedCommand;
        private readonly Helpers.RelayCommand<object> _moveToProjectCommand;
        private readonly Helpers.RelayCommand<object> _reloadProjectsCommand;

        public TabMenuItemFactory(
            Action<object> onPin,
            Action<object> onCopyName,
            Action<object> onCopyPath,
            Action<object> onOpenLocation,
            Action<object> onClose,
            Action<object> onCloseSelected,
            Action<object> onMoveToProject,
            Action<object> onReloadProjects
            ) {
            _pinCommand = new Helpers.RelayCommand<object>(onPin);
            _copyNameCommand = new Helpers.RelayCommand<object>(onCopyName);
            _copyPathCommand = new Helpers.RelayCommand<object>(onCopyPath);
            _openLocationCommand = new Helpers.RelayCommand<object>(onOpenLocation);
            _closeCommand = new Helpers.RelayCommand<object>(onClose);
            _closeSelectedCommand = new Helpers.RelayCommand<object>(onCloseSelected);
            _moveToProjectCommand = new Helpers.RelayCommand<object>(onMoveToProject);
            _reloadProjectsCommand = new Helpers.RelayCommand<object>(onReloadProjects);
        }

        public IReadOnlyList<Helpers.IMenuItem> CreateSingleSelectionContextMenu(TMEx.State.Document.TabItemBase tabItem) {
            var items = new List<Helpers.IMenuItem> {
                this.CreateCommand(State.Constants.UI.PinTab, _pinCommand, tabItem),
                this.CreateCommand(State.Constants.UI.CopyTabName, _copyNameCommand, tabItem),
                this.CreateCommand(State.Constants.UI.CopyTabPath, _copyPathCommand, tabItem),
                new Helpers.MenuItemSeparator()
            };

            if (tabItem is TMEx.State.Document.TabItemDocument) {
                items.Add(this.CreateCommand(State.Constants.UI.OpenTabLocation, _openLocationCommand, tabItem));
                items.Add(new Helpers.MenuItemSeparator());
            }

            items.Add(this.CreateCommand(State.Constants.UI.CloseTab, _closeCommand, tabItem));
            return items;
        }

        public IReadOnlyList<Helpers.IMenuItem> CreateMultipleSelectionContextMenu(
            TMEx.State.Document.TabItemBase anchorTabItem,
            IReadOnlyList<TMEx.State.Document.TabItemBase> selectedTabItems
            ) {
            return new Helpers.IMenuItem[] {
                // Popup закрывается до выполнения команды. Снимок selection не даёт клику по
                // меню изменить набор вкладок, к которым применяется команда.
                this.CreateCommand(State.Constants.UI.PinTabs, _pinCommand, selectedTabItems),
                this.CreateCommand(State.Constants.UI.CopyTabNames, _copyNameCommand, selectedTabItems),
                this.CreateCommand(State.Constants.UI.CopyTabPaths, _copyPathCommand, selectedTabItems),
                new Helpers.MenuItemSeparator(),
                this.CreateCommand(State.Constants.UI.CloseSelectedTabs, _closeSelectedCommand, anchorTabItem)
            };
        }

        public List<Helpers.IMenuItem> CreateVirtualMenu(TMEx.State.Document.TabItemDocument tabItemDocument) {
            return new List<Helpers.IMenuItem> {
                new Helpers.MenuItemHeader {
                    Header = tabItemDocument.Caption
                },
                this.CreateCommand(State.Constants.UI.CopyTabName, _copyNameCommand, tabItemDocument),
                this.CreateCommand(State.Constants.UI.CopyTabPath, _copyPathCommand, tabItemDocument),
                new Helpers.MenuItemSeparator(),
                this.CreateCommand(State.Constants.UI.OpenTabLocation, _openLocationCommand, tabItemDocument),
                this.CreateCommand(State.Constants.UI.CloseTab, _closeCommand, tabItemDocument)
            };
        }

        public Helpers.IMenuItem CreateProjectContextItem(string header, object projectContext) {
            return this.CreateCommand(header, _moveToProjectCommand, projectContext);
        }

        public Helpers.IMenuItem CreateReloadProjectsItem(TMEx.State.Document.DocumentProjectReferencesInfo referencesInfo) {
            return this.CreateCommand("Reload projects", _reloadProjectsCommand, referencesInfo);
        }

        private Helpers.MenuItemCommand CreateCommand(
            string header,
            Helpers.RelayCommand<object> command,
            object parameter
            ) {
            return new Helpers.MenuItemCommand {
                Header = header,
                Command = command,
                CommandParameterContext = parameter
            };
        }
    }
}
