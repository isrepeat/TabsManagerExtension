using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;

using Helpers.Ex;
using TMEx = TabsManagerExtension;

namespace TabsManagerExtension.Controls.Tabs {
    /// <summary>Строит меню project context и переоткрывает документы в выбранном проекте.</summary>
    internal sealed class ProjectContextController {
        private const int GroupActiveDocumentRestoreDelayMilliseconds = 450;

        private readonly EnvDTE80.DTE2 _dte;
        private readonly VirtualMenuControl _virtualMenuControl;
        private readonly TabCollectionManager _tabCollectionManager;
        private readonly Helpers.Collections.GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase> _selectionCoordinator;

        public ProjectContextController(
            EnvDTE80.DTE2 dte,
            VirtualMenuControl virtualMenuControl,
            TabCollectionManager tabCollectionManager,
            Helpers.Collections.GroupsSelectionCoordinator<TMEx.State.Document.TabItemsGroupBase, TMEx.State.Document.TabItemBase> selectionCoordinator
            ) {
            _dte = dte;
            _virtualMenuControl = virtualMenuControl;
            _tabCollectionManager = tabCollectionManager;
            _selectionCoordinator = selectionCoordinator;
        }

        public void MoveToRelatedProject(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _virtualMenuControl.HideImmediately();

            // Обычный пункт меню содержит одну связь header -> project и переключается напрямую.
            if (parameter is TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry reference) {
                this.MoveDocumentToProjectGroup(reference.DocumentEntryBase);
                return;
            }
            if (parameter is not TMEx.State.Document.DocumentProjectReferencesInfo.GroupContextEntry groupContext || !groupContext.CanSwitch) {
                return;
            }

            // Групповое переключение нескольких headers временно активирует включающие .cpp
            // и может пересоздавать document frames. Сохраняем путь, а не EnvDTE.Document:
            // старый COM-объект после retarget/reopen уже может быть недействительным.
            string? activeDocumentPath = _dte.ActiveDocument?.FullName;
            // План подбирает минимальный практически достижимый набор translation units,
            // способный переключить контекст всех выбранных headers.
            var contextSwitchPlan = this.BuildGroupContextSwitchPlan(groupContext.DocumentReferences);
            // Один .cpp может включать несколько выбранных headers. Набор позволяет нижнему
            // уровню повторно использовать уже активированный source вместо лишнего открытия.
            var activatedSources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var documentReference in groupContext.DocumentReferences.ToList()) {
                contextSwitchPlan.TryGetValue(documentReference, out var sourcePath);
                this.MoveDocumentToProjectGroup(
                    documentReference.DocumentEntryBase,
                    playFeedback: false,
                    preferredContextSwitchSourcePath: sourcePath,
                    activatedContextSwitchSources: activatedSources
                );
            }

            //Console.Beep(frequency: 1000, duration: 300);
            // SharedItem.OpenWithProjectContext использует fire-and-forget продолжение: через
            // 300 мс header переоткрывается в новой hierarchy после обработки translation unit
            // C++ language service. Дополнительный запас даёт DTE завершить DocumentOpened и
            // WindowActivated; продолжения на 100 и 20 мс укладываются в тот же интервал.
            // Простой Dispatcher yield здесь недостаточен: он выполнится раньше 300-мс таймера
            // и последующее переоткрытие снова отберёт активный frame.
            // Эту задержку можно удалить после перевода всей цепочки OpenWithProjectContext
            // с fire-and-forget void на Task: тогда здесь следует дождаться всех операций через
            // Task.WhenAll и восстановить исходный документ по фактическому завершению reopen.
            VsixThreadHelper.RunOnUiThread(async () => {
                await Task.Delay(GroupActiveDocumentRestoreDelayMilliseconds);

                var activeDocument = _dte.Documents
                    .Cast<EnvDTE.Document>()
                    .FirstOrDefault(document => string.Equals(
                        document.FullName,
                        activeDocumentPath,
                        StringComparison.OrdinalIgnoreCase
                    ));

                if (activeDocument != null) {
                    activeDocument.Activate();
                }
                else if (!string.IsNullOrEmpty(activeDocumentPath)) {
                    Helpers.Diagnostic.Logger.LogDebug($"[ProjectContextController] Cannot restore active document '{activeDocumentPath}' because its current frame was not found.");
                }
            });
        }

        public void MoveToRelatedProjectFile(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();
            _virtualMenuControl.HideImmediately();

            if (parameter is not ProjectContextSourceEntry sourceEntry) {
                return;
            }

            var projectEntry = sourceEntry.ProjectContext switch {
                TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry reference => reference.ProjectEntry,
                TMEx.State.Document.DocumentProjectReferencesInfo.GroupContextEntry groupContext => groupContext.ProjectEntry,
                _ => null
            };
            if (projectEntry?.MultiState.Current is not VsShell.Project.LoadedProject loadedProject) {
                return;
            }

            // Один и тот же физический .cpp может иметь несколько project representations.
            // Клик по hierarchy item целевого проекта гарантирует правильный project context,
            // чего не гарантирует простое ItemOperations.OpenFile(path).
            var sourceDocument = VsShell.Solution.Services.SolutionHierarchyAnalyzerService.Instance
                .SourcesRepresentationsTable
                .GetDocumentByProjectAndDocumentPath(projectEntry.BaseViewModel, sourceEntry.SourcePath);
            if (sourceDocument?.BaseViewModel.HierarchyItemEntry.BaseViewModel is VsShell.Hierarchy.RealHierarchyItem hierarchyItem) {
                int result = VsShell.Utils.VsHierarchyUtils.ClickOnSolutionHierarchyItem(
                    loadedProject.ProjectHierarchy.VsRealHierarchy,
                    hierarchyItem.ItemId
                );

                ErrorHandler.ThrowOnFailure(result);
                return;
            }

            // Representation может отсутствовать во временно неполной модели solution.
            // В таком случае оставляем полезный fallback: открыть существующий файл как текст.
            if (!File.Exists(sourceEntry.SourcePath)) {
                Helpers.Diagnostic.Logger.LogWarning($"[ProjectContextController] File '{sourceEntry.SourcePath}' does not exist.");
                return;
            }

            _dte.ItemOperations
                .OpenFile(sourceEntry.SourcePath, EnvDTE.Constants.vsViewKindTextView)
                .Activate();
        }

        public void ToggleIncludersMenu(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (parameter is not FrameworkElement anchor ||
                anchor.DataContext is not Helpers.MenuItemCommand menuItem ||
                menuItem.CommandParameterContext is not object projectContext) {

                return;
            }

            // DataContext используется как identity конкретного project-context пункта.
            // Повторный клик по тому же пункту закрывает child menu, а не перестраивает его.
            if (_virtualMenuControl.IsChildMenuOpen &&
                ReferenceEquals(_virtualMenuControl.CurrentChildMenuDataContext, projectContext)) {

                _virtualMenuControl.HideChild();
                return;
            }

            this.ShowIncludersMenu(anchor, projectContext);
        }

        public void HandleMenuItemMouseEnter(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Hover переключает содержимое только уже открытого child menu. Сам по себе
            // MouseEnter не должен неожиданно раскрывать подменю пользователю.
            if (!_virtualMenuControl.IsChildMenuOpen ||
                parameter is not MenuItem menuItem ||
                menuItem.DataContext is not Helpers.MenuItemCommand menuItemCommand) {

                return;
            }

            object projectContext = menuItemCommand.CommandParameterContext;
            if (!ReferenceEquals(_virtualMenuControl.CurrentChildMenuDataContext, projectContext)) {
                this.ShowIncludersMenu(menuItem, projectContext);
            }
        }

        public void AppendProjectMenuItems(
            List<Helpers.IMenuItem> menuItems,
            TMEx.State.Document.TabItemDocument anchorTabItem,
            TabMenuItemFactory menuItemFactory
            ) {
            // Специальный групповой режим включается только для мультивыбора headers, в который
            // входит anchor контекстного меню. Для остальных случаев строится обычное меню вкладки.
            if (this.TryGetSelectedHeaderGroup(anchorTabItem, out var selectedHeaders)) {
                var groupContexts = this.BuildGroupContextEntries(selectedHeaders);
                if (groupContexts.Count == 0) {
                    return;
                }

                menuItems.Add(new Helpers.MenuItemSeparator());
                // Сначала показываем проекты, доступные всем headers и допускающие групповое
                // переключение. Частичные проекты отделяем: они полезны как информация,
                // но GroupContextEntry.CanSwitch не позволит применить неполную операцию.
                var commonContexts = groupContexts.Where(context => context.IsAvailableForAll).ToList();
                var differingContexts = groupContexts.Where(context => !context.IsAvailableForAll).ToList();
                this.AppendContextItems(menuItems, commonContexts, menuItemFactory);
                if (commonContexts.Count > 0 && differingContexts.Count > 0) {
                    menuItems.Add(new Helpers.MenuItemSeparator());
                }

                this.AppendContextItems(menuItems, differingContexts, menuItemFactory);
                return;
            }

            // Для одиночной вкладки GetAvailableReferences намеренно скрывает единственный
            // собственный проект: переключать контекст имеет смысл только при наличии выбора.
            var references = anchorTabItem.DocumentProjectReferencesInfo.GetAvailableReferences();
            if (references.Count == 0) {
                return;
            }

            menuItems.Add(new Helpers.MenuItemSeparator());
            foreach (var reference in references) {
                menuItems.Add(menuItemFactory.CreateProjectContextItem(
                    reference.ProjectEntry.BaseViewModel.Name,
                    reference
                ));
            }

            // Пункт reload нужен, когда часть ссылок относится к unloaded project и обычное
            // переключение пока невозможно.
            menuItems.Add(menuItemFactory.CreateReloadProjectsItem(anchorTabItem.DocumentProjectReferencesInfo));
        }

        public void ReloadProjects(object parameter) {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (parameter is not TMEx.State.Document.DocumentProjectReferencesInfo referencesInfo) {
                return;
            }

            // Работаем по текущему снимку references и пропускаем уже загруженные проекты.
            // После reload анализатор solution асинхронно перестроит сами RefEntry.
            foreach (var reference in referencesInfo.References) {
                var projectViewModel = reference.ProjectEntry.BaseViewModel;
                if (projectViewModel is VsShell.Project.UnloadedProject) {
                    VsShell.Utils.VsHierarchyUtils.ReloadProject(projectViewModel.ProjectGuid);
                }
            }
        }

        private void MoveDocumentToProjectGroup(
            VsShell.Document.DocumentEntryBase documentEntry,
            bool playFeedback = true,
            string? preferredContextSwitchSourcePath = null,
            ISet<string>? activatedContextSwitchSources = null
            ) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Сначала меняется реальный project context в VS. Операция может retarget-нуть
            // существующий frame либо закрыть и открыть документ заново.
            OpenWithProjectContext(
                documentEntry,
                playFeedback,
                preferredContextSwitchSourcePath,
                activatedContextSwitchSources
            );

            // После VS-операции ищем TabItem по стабильному пути. Автоматическая классификация
            // не знает выбранный пользователем target project, поэтому явно переносим вкладку
            // в default-группу этого проекта.
            var documentViewModel = documentEntry.BaseViewModel;
            string filePath = documentViewModel.HierarchyItemEntry.BaseViewModel.FilePath;
            var tabItemDocument = _tabCollectionManager.Find(filePath);
            if (tabItemDocument == null) {
                return;
            }

            _tabCollectionManager.Remove(tabItemDocument);
            _tabCollectionManager.AddToGroup(
                tabItemDocument,
                new TMEx.State.Document.TabItemsDefaultGroup(documentViewModel.ProjectBaseViewModel.Name)
            );
        }

        private void ShowIncludersMenu(FrameworkElement anchor, object projectContext) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Child menu показывает только .cpp translation units целевого проекта, через
            // которые language service может установить context выбранного header.
            var sourcePaths = GetProjectContextSwitchSourcePaths(projectContext);
            string projectName = projectContext switch {
                TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry reference => reference.ProjectEntry.BaseViewModel.Name,
                TMEx.State.Document.DocumentProjectReferencesInfo.GroupContextEntry groupContext => groupContext.ProjectEntry.BaseViewModel.Name,
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
                    // Команде нужны одновременно исходный project context и конкретный source;
                    // отдельный DTO не даёт потерять один из параметров в MenuItem.DataContext.
                    childItems.Add(new Helpers.MenuItemCommand {
                        Header = Path.GetFileName(sourcePath),
                        Command = new Helpers.RelayCommand<object>(this.MoveToRelatedProjectFile),
                        CommandParameterContext = new ProjectContextSourceEntry(projectContext, sourcePath)
                    });
                }
            }

            // VirtualMenuControl не участвует в обычном WPF placement. Вычисляем экранную
            // позицию с учётом DPI и открываем child справа от пункта-источника.
            Point screenPoint = anchor.ex_ToDpiAwareScreen(new Point(anchor.ActualWidth + 8, 0));
            _virtualMenuControl.ShowChild(screenPoint, projectContext, childItems);
        }

        private bool TryGetSelectedHeaderGroup(
            TMEx.State.Document.TabItemDocument anchorTabItem,
            out IReadOnlyList<TMEx.State.Document.TabItemDocument> selectedHeaders
            ) {
            // Групповая команда допустима, только если anchor принадлежит selection и каждый
            // выбранный элемент — header. Так контекстное меню не применит действие к скрытой
            // части selection или к .cpp/tool window, для которых этот сценарий бессмыслен.
            var selectedItems = _selectionCoordinator.SelectedItems.Select(entry => entry.Item).ToList();
            bool anchorIsSelected = selectedItems.Any(item => ReferenceEquals(item, anchorTabItem));
            bool allSelectedItemsAreHeaders = selectedItems.All(
                item => item is TMEx.State.Document.TabItemDocument document && IsHeaderPath(document.FullName)
            );

            if (_selectionCoordinator.SelectionState != Helpers.Enums.SelectionState.Multiple ||
                !anchorIsSelected ||
                selectedItems.Count < 2 ||
                !allSelectedItemsAreHeaders) {

                selectedHeaders = Array.Empty<TMEx.State.Document.TabItemDocument>();
                return false;
            }

            selectedHeaders = selectedItems.Cast<TMEx.State.Document.TabItemDocument>().ToList();
            return true;
        }

        private IReadOnlyList<TMEx.State.Document.DocumentProjectReferencesInfo.GroupContextEntry> BuildGroupContextEntries(
            IReadOnlyList<TMEx.State.Document.TabItemDocument> selectedHeaders
            ) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Здесь includeSingleProject=true принципиален: проект, единственный для одного
            // header, всё равно может быть общим target project для всего мультивыбора.
            var referencesByDocument = selectedHeaders
                .Select(header => header.DocumentProjectReferencesInfo
                    .GetAvailableReferences(includeSingleProject: true)
                    .GroupBy(reference => reference.ProjectEntry.BaseViewModel.ProjectGuid)
                    .ToDictionary(group => group.Key, group => group.First()))
                .ToList();
            // Берём объединение project GUID, а затем для каждого проекта собираем только те
            // document references, которые реально существуют у выбранных headers.
            var projectGuids = referencesByDocument.SelectMany(references => references.Keys).Distinct().ToList();
            var result = new List<TMEx.State.Document.DocumentProjectReferencesInfo.GroupContextEntry>();
            foreach (var projectGuid in projectGuids) {
                var documentReferences = referencesByDocument
                    .Where(references => references.ContainsKey(projectGuid))
                    .Select(references => references[projectGuid])
                    .ToList();

                result.Add(new TMEx.State.Document.DocumentProjectReferencesInfo.GroupContextEntry(
                    documentReferences[0].ProjectEntry,
                    documentReferences,
                    documentReferences.Count == selectedHeaders.Count
                ));
            }

            // Полностью общие контексты идут первыми; стабильная сортировка по UniqueName
            // сохраняет предсказуемый порядок при одинаковых отображаемых именах проектов.
            return result
                .OrderByDescending(context => context.IsAvailableForAll)
                .ThenBy(context => context.ProjectEntry.BaseViewModel.UniqueName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private IReadOnlyDictionary<TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry, string> BuildGroupContextSwitchPlan(
            IReadOnlyList<TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry> documentReferences
            ) {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Для каждого header строим множество транзитивно включающих .cpp в target project.
            var candidatesByDocument = documentReferences.ToDictionary(
                reference => reference,
                reference => GetProjectContextSwitchSourcePaths(reference).ToHashSet(StringComparer.OrdinalIgnoreCase)
            );
            var openDocumentPaths = _dte.Documents
                .Cast<EnvDTE.Document>()
                .Select(document => document.FullName)
                .Where(path => !string.IsNullOrEmpty(path))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            // Headers без кандидатов не участвуют в set-cover цикле: для них ниже будет warning,
            // а OpenWithProjectContext получит null и применит собственный fallback.
            var uncoveredDocuments = candidatesByDocument
                .Where(pair => pair.Value.Count > 0)
                .Select(pair => pair.Key)
                .ToHashSet();
            var result = new Dictionary<TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry, string>();

            while (uncoveredDocuments.Count > 0) {
                // Жадный set cover выбирает .cpp, покрывающий максимум ещё не обработанных
                // headers. При равенстве предпочитается закрытый source: нижний уровень может
                // кратко открыть и закрыть его, почти не затрагивая видимые editor frames.
                string selectedSourcePath = uncoveredDocuments
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

                // Все покрытые headers получают один preferred source; activatedSources из
                // вызывающего метода не даст физически активировать его повторно для каждого.
                var coveredDocuments = uncoveredDocuments
                    .Where(reference => candidatesByDocument[reference].Contains(selectedSourcePath))
                    .ToList();
                foreach (var coveredDocument in coveredDocuments) {
                    result[coveredDocument] = selectedSourcePath;
                    uncoveredDocuments.Remove(coveredDocument);
                }
            }

            foreach (var reference in candidatesByDocument.Where(pair => pair.Value.Count == 0)) {
                Helpers.Diagnostic.Logger.LogWarning($"[ProjectContextController] No transitive .cpp was found for '{reference.Key.DocumentEntryBase.BaseViewModel.HierarchyItemEntry.BaseViewModel.FilePath}'.");
            }

            return result;
        }

        private static IReadOnlyList<string> GetProjectContextSwitchSourcePaths(object projectContext) {
            // Child menu принимает и одиночный, и групповой context. Нормализуем оба варианта
            // в набор references, объединяем пути и сортируем сначала по короткому имени файла.
            IEnumerable<TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry> references = projectContext switch {
                TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry reference => new[] { reference },
                TMEx.State.Document.DocumentProjectReferencesInfo.GroupContextEntry groupContext => groupContext.DocumentReferences,
                _ => Array.Empty<TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry>()
            };

            return references
                .SelectMany(GetProjectContextSwitchSourcePaths)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
                .ThenBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static IReadOnlyList<string> GetProjectContextSwitchSourcePaths(
            TMEx.State.Document.DocumentProjectReferencesInfo.RefEntry reference
            ) {
            // RefEntry описывает representation в target project, но поиск includers реализован
            // текущим Document multi-state. Оба объекта должны быть валидны и project загружен.
            var document = GetCurrentProjectContextDocument(reference.DocumentEntryBase);
            var targetProject = reference.ProjectEntry.MultiState.Current as VsShell.Project.LoadedProject;
            return document != null && targetProject != null
                ? document.GetProjectContextSwitchSourcePaths(targetProject)
                : Array.Empty<string>();
        }

        private static VsShell.Document.Document? GetCurrentProjectContextDocument(
            VsShell.Document.DocumentEntryBase entry
            ) {
            // Только ExternalInclude и SharedItem предоставляют Document-логику переключения.
            // Invalidated/unknown состояния намеренно возвращают null и блокируют построение меню.
            return entry switch {
                VsShell.Document.ExternalIncludeEntry externalIncludeEntry =>
                    externalIncludeEntry.MultiState.Current as VsShell.Document.Document,
                VsShell.Document.SharedItemEntry sharedItemEntry =>
                    sharedItemEntry.MultiState.Current as VsShell.Document.Document,
                _ => null
            };
        }

        private static bool IsHeaderPath(string path) {
            string extension = Path.GetExtension(path);
            return string.Equals(extension, ".h", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".hpp", StringComparison.OrdinalIgnoreCase);
        }

        private static void OpenWithProjectContext(
            VsShell.Document.DocumentEntryBase documentEntry,
            bool playFeedback,
            string? preferredContextSwitchSourcePath,
            ISet<string>? activatedContextSwitchSources
            ) {
            // External include и shared item требуют разных VS-механизмов, но оба принимают
            // preferred .cpp и общий набор уже активированных sources для групповой операции.
            if (documentEntry is VsShell.Document.ExternalIncludeEntry externalIncludeEntry) {
                if (externalIncludeEntry.MultiState.Current is VsShell.Document.ExternalInclude externalInclude) {
                    externalInclude.OpenWithProjectContext(
                        preferredContextSwitchSourcePath,
                        activatedContextSwitchSources
                    );
                    if (playFeedback) {
                        Console.Beep(frequency: 1000, duration: 300);
                    }
                }
                else if (externalIncludeEntry.MultiState.Current is VsShell.Document.InvalidatedDocument) {
                    // Для invalidated external include корректного fallback пока нет: breakpoint
                    // сохраняет прежнее диагностическое поведение в отладочной сборке.
                    System.Diagnostics.Debugger.Break();
                }

                return;
            }

            if (documentEntry is VsShell.Document.SharedItemEntry sharedItemEntry) {
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
                    // Invalidated shared item умеет сам заново разрешить актуальную representation,
                    // поэтому здесь можно делегировать открытие без preferred source.
                    invalidatedDocument.OpenWithProjectContext();
                }
            }
        }

        private void AppendContextItems(
            ICollection<Helpers.IMenuItem> menuItems,
            IEnumerable<TMEx.State.Document.DocumentProjectReferencesInfo.GroupContextEntry> contexts,
            TabMenuItemFactory menuItemFactory
            ) {
            // GroupContextEntry уже содержит CanSwitch и полный набор document references;
            // фабрика превращает его в команду и визуально отражает недоступное состояние.
            foreach (var context in contexts) {
                menuItems.Add(menuItemFactory.CreateProjectContextItem(
                    context.ProjectEntry.BaseViewModel.Name,
                    context
                ));
            }
        }

        // Параметр команды child menu: target project context плюс выбранный пользователем
        // translation unit, который нужно открыть именно в hierarchy этого проекта.
        private sealed class ProjectContextSourceEntry {
            public object ProjectContext { get; }
            public string SourcePath { get; }

            public ProjectContextSourceEntry(object projectContext, string sourcePath) {
                this.ProjectContext = projectContext;
                this.SourcePath = sourcePath;
            }
        }
    }
}
