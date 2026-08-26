using System;
using System.IO;
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

        public TabRenameResult Rename(TMEx.State.Document.TabItemDocument tabItemDocument, string? proposedName) {
            ThreadHelper.ThrowIfNotOnUIThread();

            string oldPath = tabItemDocument.FullName;
            string oldExtension = Path.GetExtension(oldPath);
            string newName = proposedName?.Trim() ?? string.Empty;
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
            catch (Exception ex) {
                Helpers.Diagnostic.Logger.LogError($"Failed to rename tab '{oldPath}' to '{newName}': {ex}");
                return TabRenameResult.Failure($"Visual Studio could not rename '{Path.GetFileName(oldPath)}' to '{newName}'.\n\n{ex.Message}");
            }
        }
    }
}
