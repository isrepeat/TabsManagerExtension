using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Microsoft.VisualStudio.Shell;

using TMEx = TabsManagerExtension;


namespace TabsManagerExtension.Controls.Tabs {
    internal sealed class TabRenameResult {
        public bool Succeeded { get; }
        public string? ErrorMessage { get; }

        private TabRenameResult(bool succeeded, string? errorMessage) {
            this.Succeeded = succeeded;
            this.ErrorMessage = errorMessage;
        }

        public static TabRenameResult Success() {
            return new TabRenameResult(true, null);
        }

        public static TabRenameResult Failure(string errorMessage) {
            return new TabRenameResult(false, errorMessage);
        }
    }

    /// <summary>Проверяет новое имя и переименовывает файл через project item или файловую систему.</summary>
    internal sealed class TabRenameService {
        private static readonly HashSet<string> RenamableExtensions = new(StringComparer.OrdinalIgnoreCase) {
            ".c",
            ".cpp",
            ".h",
            ".cxx",
            ".hxx"
        };

        // Возвращает выделенную группу одноимённых файлов с разными поддерживаемыми расширениями,
        // которую можно массово переименовать через шаблон "новое_имя.*".
        public static IReadOnlyList<TMEx.State.Document.TabItemDocument> GetSelectedRenameGroup(
            TMEx.State.Document.TabItemDocument tabItemDocument,
            IEnumerable<TMEx.State.Document.TabItemBase> selectedTabItems
            ) {
            var selectedItems = selectedTabItems.ToList();
            var documentTabItems = selectedItems
                .OfType<TMEx.State.Document.TabItemDocument>()
                .ToList();
            if (documentTabItems.Count < 2 ||
                documentTabItems.Count != selectedItems.Count ||
                !documentTabItems.Any(item => ReferenceEquals(item, tabItemDocument))) {
                return Array.Empty<TMEx.State.Document.TabItemDocument>();
            }

            string directory = Path.GetDirectoryName(tabItemDocument.FullName) ?? string.Empty;
            string baseName = Path.GetFileNameWithoutExtension(tabItemDocument.FullName);
            bool isRenameGroup = documentTabItems.All(candidate =>
                string.Equals(Path.GetDirectoryName(candidate.FullName) ?? string.Empty, directory, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(Path.GetFileNameWithoutExtension(candidate.FullName), baseName, StringComparison.OrdinalIgnoreCase) &&
                RenamableExtensions.Contains(Path.GetExtension(candidate.FullName))
            );
            bool haveDifferentExtensions = documentTabItems
                .Select(candidate => Path.GetExtension(candidate.FullName))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count() == documentTabItems.Count;
            return isRenameGroup && haveDifferentExtensions
                ? documentTabItems
                : Array.Empty<TMEx.State.Document.TabItemDocument>();
        }

        public TabRenameResult Rename(
            TMEx.State.Document.TabItemDocument tabItemDocument,
            string? proposedName,
            IReadOnlyList<TMEx.State.Document.TabItemDocument> renameGroupTabItems
            ) {
            ThreadHelper.ThrowIfNotOnUIThread();

            string oldPath = tabItemDocument.FullName;
            string oldExtension = Path.GetExtension(oldPath);
            string newName = proposedName?.Trim() ?? string.Empty;
            if (string.Equals(Path.GetExtension(newName), ".*", StringComparison.Ordinal)) {
                return this.RenameGroupWithPreservedExtensions(tabItemDocument, newName, renameGroupTabItems);
            }

            if (string.Equals(Path.GetFileName(oldPath), newName, StringComparison.Ordinal)) {
                return TabRenameResult.Success();
            }

            if (!RenamableExtensions.Contains(oldExtension)) {
                return TabRenameResult.Failure($"Files with the '{oldExtension}' extension cannot be renamed from Tabs Manager.");
            }
            if (string.IsNullOrWhiteSpace(newName) ||
                !string.Equals(newName, Path.GetFileName(newName), StringComparison.Ordinal) ||
                newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {

                return TabRenameResult.Failure("Enter a valid file name without a directory path.");
            }

            string newExtension = Path.GetExtension(newName);
            if (!RenamableExtensions.Contains(newExtension)) {
                return TabRenameResult.Failure($"The target extension '{newExtension}' is not supported. Use .c, .cpp, .h, .cxx, or .hxx.");
            }

            string? directory = Path.GetDirectoryName(oldPath);
            string newPath = directory == null ? newName : Path.Combine(directory, newName);
            if (!string.Equals(oldPath, newPath, StringComparison.OrdinalIgnoreCase) && File.Exists(newPath)) {
                return TabRenameResult.Failure($"A file named '{newName}' already exists in this directory.");
            }

            EnvDTE.ProjectItem? projectItem;
            try {
                projectItem = tabItemDocument.ShellDocument.Document.ProjectItem;
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to resolve ProjectItem for '{oldPath}': {ex}");
                return TabRenameResult.Failure("Visual Studio could not resolve the project item for this document.");
            }

            if (projectItem == null) {
                return TabRenameResult.Failure("This document is not represented by a project item and cannot be safely renamed.");
            }

            try {
                // ProjectItem.Name запускает штатное переименование project system: обновляются файл,
                // открытый document moniker и ссылки проекта.
                projectItem.Name = newName;
                tabItemDocument.Caption = tabItemDocument.ShellDocument.Document.Name;
                tabItemDocument.FullName = tabItemDocument.ShellDocument.Document.FullName;

                return TabRenameResult.Success();
            }
            catch (NotImplementedException ex) {
                // Miscellaneous Files не реализует ProjectItem.Name. Закрываем документ до
                // файлового rename, затем открываем его по новому пути, чтобы VS получил новый moniker.
                Helpers.Diagnostic.Logger.LogWarning($"ProjectItem rename is not implemented for '{oldPath}': {ex.Message}");
                return this.RenameMiscellaneousFile(tabItemDocument, oldPath, newPath, newName);
            }
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to rename tab '{oldPath}' to '{newName}': {ex}");
                return TabRenameResult.Failure($"Visual Studio could not rename '{Path.GetFileName(oldPath)}' to '{newName}'.\n\n{ex.Message}");
            }
        }

        // Переименовывает все файлы согласованной группы, заменяя только базовое имя и сохраняя
        // индивидуальное расширение каждого файла.
        private TabRenameResult RenameGroupWithPreservedExtensions(
            TMEx.State.Document.TabItemDocument tabItemDocument,
            string proposedName,
            IReadOnlyList<TMEx.State.Document.TabItemDocument> renameGroupTabItems
            ) {
            string newBaseName = Path.GetFileNameWithoutExtension(proposedName);
            if (string.IsNullOrWhiteSpace(newBaseName) ||
                !string.Equals(newBaseName, Path.GetFileName(newBaseName), StringComparison.Ordinal) ||
                newBaseName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {

                return TabRenameResult.Failure("Enter a valid file name before '.*'.");
            }

            var tabItems = renameGroupTabItems.Count > 1
                ? renameGroupTabItems
                : Array.Empty<TMEx.State.Document.TabItemDocument>();
            if (tabItems.Count == 0 || !tabItems.Any(item => ReferenceEquals(item, tabItemDocument))) {
                return TabRenameResult.Failure("Use '.*' only for a selected group of files with the same name.");
            }

            foreach (var item in tabItems) {
                string directory = Path.GetDirectoryName(item.FullName) ?? string.Empty;
                string targetName = newBaseName + Path.GetExtension(item.FullName);
                string targetPath = Path.Combine(directory, targetName);
                if (!string.Equals(item.FullName, targetPath, StringComparison.OrdinalIgnoreCase) && File.Exists(targetPath)) {
                    return TabRenameResult.Failure($"A file named '{targetName}' already exists in this directory.");
                }
            }

            foreach (var item in tabItems) {
                string targetName = newBaseName + Path.GetExtension(item.FullName);
                var result = this.Rename(item, targetName, Array.Empty<TMEx.State.Document.TabItemDocument>());
                if (!result.Succeeded) {
                    return TabRenameResult.Failure(
                        $"Some files were renamed before '{Path.GetFileName(item.FullName)}' failed.\n\n{result.ErrorMessage}"
                    );
                }
            }

            return TabRenameResult.Success();
        }

        // Обрабатывает файлы из Miscellaneous Files: у них нет реализации ProjectItem.Name,
        // поэтому файл переименовывается через файловую систему и заново открывается в VS.
        private TabRenameResult RenameMiscellaneousFile(
            TMEx.State.Document.TabItemDocument tabItemDocument,
            string oldPath,
            string newPath,
            string newName
            ) {
            bool fileMoved = false;
            try {
                var document = tabItemDocument.ShellDocument.Document;
                document.Save();
                document.Close(EnvDTE.vsSaveChanges.vsSaveChangesNo);
                File.Move(oldPath, newPath);
                fileMoved = true;
                PackageServices.Dte2.ItemOperations.OpenFile(newPath);
                return TabRenameResult.Success();
            }
            catch (Exception ex) {
                if (!fileMoved && File.Exists(oldPath)) {
                    try {
                        PackageServices.Dte2.ItemOperations.OpenFile(oldPath);
                    }
                    catch (Exception reopenException) {
                        Helpers.Diagnostic.Logger.LogWarning($"Failed to reopen '{oldPath}' after rename failure: {reopenException.Message}");
                    }
                }

                Helpers.Diagnostic.Logger.LogError($"Failed to rename miscellaneous file '{oldPath}' to '{newPath}': {ex}");
                return TabRenameResult.Failure($"Visual Studio could not rename '{Path.GetFileName(oldPath)}' to '{newName}'.\n\n{ex.Message}");
            }
        }
    }
}
