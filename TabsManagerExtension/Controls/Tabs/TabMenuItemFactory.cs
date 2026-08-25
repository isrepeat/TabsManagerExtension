using System;
using System.Collections.Generic;

using TabsManagerExtension.State.Document;


namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Формирует пункты context и virtual menu из команд вкладок.</summary>
    internal sealed class TabMenuItemFactory {
        private readonly Helpers.RelayCommand<object> _copyNameCommand;
        private readonly Helpers.RelayCommand<object> _copyPathCommand;
        private readonly Helpers.RelayCommand<object> _openLocationCommand;
        private readonly Helpers.RelayCommand<object> _closeCommand;
        private readonly Helpers.RelayCommand<object> _closeSelectedCommand;
        private readonly Helpers.RelayCommand<object> _moveToProjectCommand;
        private readonly Helpers.RelayCommand<object> _reloadProjectsCommand;

        public TabMenuItemFactory(
            Action<object> copyName,
            Action<object> copyPath,
            Action<object> openLocation,
            Action<object> close,
            Action<object> closeSelected,
            Action<object> moveToProject,
            Action<object> reloadProjects
            ) {
            _copyNameCommand = new Helpers.RelayCommand<object>(copyName);
            _copyPathCommand = new Helpers.RelayCommand<object>(copyPath);
            _openLocationCommand = new Helpers.RelayCommand<object>(openLocation);
            _closeCommand = new Helpers.RelayCommand<object>(close);
            _closeSelectedCommand = new Helpers.RelayCommand<object>(closeSelected);
            _moveToProjectCommand = new Helpers.RelayCommand<object>(moveToProject);
            _reloadProjectsCommand = new Helpers.RelayCommand<object>(reloadProjects);
        }

        public IReadOnlyList<Helpers.IMenuItem> CreateSingleSelectionContextMenu(TabItemBase tabItem) {
            var items = new List<Helpers.IMenuItem> {
                this.CreateCommand(State.Constants.UI.CopyTabName, _copyNameCommand, tabItem),
                this.CreateCommand(State.Constants.UI.CopyTabPath, _copyPathCommand, tabItem),
                new Helpers.MenuItemSeparator()
            };

            if (tabItem is TabItemDocument) {
                items.Add(this.CreateCommand(State.Constants.UI.OpenTabLocation, _openLocationCommand, tabItem));
                items.Add(new Helpers.MenuItemSeparator());
            }

            items.Add(this.CreateCommand(State.Constants.UI.CloseTab, _closeCommand, tabItem));
            return items;
        }

        public IReadOnlyList<Helpers.IMenuItem> CreateMultipleSelectionContextMenu(TabItemBase anchorTabItem) {
            return new[] {
                this.CreateCommand(State.Constants.UI.CloseSelectedTabs, _closeSelectedCommand, anchorTabItem)
            };
        }

        public List<Helpers.IMenuItem> CreateVirtualMenu(TabItemDocument tabItemDocument) {
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

        public Helpers.IMenuItem CreateReloadProjectsItem(DocumentProjectReferencesInfo referencesInfo) {
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
