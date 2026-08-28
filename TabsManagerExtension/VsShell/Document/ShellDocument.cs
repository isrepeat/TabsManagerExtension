using System;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace TabsManagerExtension.VsShell.Document {
    public class ShellDocument {
        public EnvDTE.Document Document { get; private set; }

        public ShellDocument(EnvDTE.Document document) {
            ThreadHelper.ThrowIfNotOnUIThread();

            this.Document = document;
        }


        public string GetDocumentProjectName() {
            // Для shared/external header VS может не создать DTE.ProjectItem. Однако frame
            // сохраняет hierarchy проекта, из которого IDE разрешила #include. Проверяем его
            // первым: у одного физического header может быть несколько project context-ов,
            // и именно context открытия определяет группу вкладки.
            var windowFrameProjectName = this.GetProjectNameFromWindowFrame();
            if (windowFrameProjectName != null) {
                return windowFrameProjectName;
            }

            try {
                var projectItem = this.Document.ProjectItem?.ContainingProject;
                if (projectItem != null) {
                    return projectItem.Name;
                }
            }
            catch {
                // External Dependencies нередко не предоставляет ProjectItem через DTE.
            }

            return "Без проекта";
        }


        public void Close(string documentPath) {
            ThreadHelper.ThrowIfNotOnUIThread();

            EnvDTE.Document? currentDocument = null;
            foreach (EnvDTE.Document document in PackageServices.Dte2.Documents) {
                try {
                    if (string.Equals(document.FullName, documentPath, StringComparison.OrdinalIgnoreCase)) {
                        currentDocument = document;
                        break;
                    }
                }
                catch {
                    // В Documents могут ещё находиться устаревшие COM-объекты временных файлов.
                }
            }

            if (currentDocument != null) {
                // DTE координирует последовательное закрытие документов. Прямой CloseFrame при
                // закрытии пачки может оставить часть уже скрытых фреймов незакрытыми, и VS
                // повторно покажет их при следующей активации документа.
                currentDocument.Close(EnvDTE.vsSaveChanges.vsSaveChangesPrompt);
                return;
            }

            // Fallback нужен для устаревшей DTE-обёртки, которая ещё не исчезла из модели вкладки.
            this.Document.Close(EnvDTE.vsSaveChanges.vsSaveChangesPrompt);
        }

      

        public bool IsDocumentInPreviewTab() {
            ThreadHelper.ThrowIfNotOnUIThread();

            var vsUiShell = PackageServices.VsUIShell;
            if (vsUiShell == null)
                return false;

            vsUiShell.GetDocumentWindowEnum(out IEnumWindowFrames windowFramesEnum);
            IVsWindowFrame[] frameArray = new IVsWindowFrame[1];
            uint fetched;

            while (windowFramesEnum.Next(1, frameArray, out fetched) == VSConstants.S_OK && fetched == 1) {
                IVsWindowFrame frame = frameArray[0];
                if (frame == null)
                    continue;

                // Получаем путь к документу
                if (ErrorHandler.Succeeded(frame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out object docPathObj)) &&
                    docPathObj is string docPath &&
                    string.Equals(docPath, this.Document.FullName, StringComparison.OrdinalIgnoreCase)) {
                    // Проверяем, является ли окно временным (предварительный просмотр)
                    if (ErrorHandler.Succeeded(frame.GetProperty((int)__VSFPROPID5.VSFPROPID_IsProvisional, out object isProvisionalObj)) &&
                        isProvisionalObj is bool isProvisional) {
                        return isProvisional;
                    }
                }
            }

            return false;
        }


        public void OpenDocumentAsPinned() {
            ThreadHelper.ThrowIfNotOnUIThread();

            var vsUIShellOpenDocument = PackageServices.VsUIShellOpenDocument;
            if (vsUIShellOpenDocument == null) {
                return;
            }

            Guid logicalView = VSConstants.LOGVIEWID_Primary;
            IVsUIHierarchy hierarchy;
            uint itemId;
            IVsWindowFrame windowFrame;
            Microsoft.VisualStudio.OLE.Interop.IServiceProvider serviceProvider;

            // Повторное открытие документа
            int hr = vsUIShellOpenDocument.OpenDocumentViaProject(
                this.Document.FullName,
                ref logicalView,
                out serviceProvider,
                out hierarchy,
                out itemId,
                out windowFrame);

            if (ErrorHandler.Succeeded(hr) && windowFrame != null) {
                windowFrame.Show();
            }
        }


        //
        // Internal logic
        //
        private string? GetProjectNameFromWindowFrame() {
            ThreadHelper.ThrowIfNotOnUIThread();

            var vsUiShell = PackageServices.VsUIShell;
            // Document window enum содержит все открытые editor frames, включая External
            // Dependencies, для которых DTE.Document.ProjectItem возвращает null.
            if (vsUiShell == null ||
                ErrorHandler.Failed(vsUiShell.GetDocumentWindowEnum(out var windowFramesEnum))) {

                return null;
            }

            var frameArray = new IVsWindowFrame[1];
            while (windowFramesEnum.Next(1, frameArray, out var fetched) == VSConstants.S_OK && fetched == 1) {
                var frame = frameArray[0];
                // Путь — стабильный идентификатор документа. Frame caption использовать нельзя:
                // одинаковые имена header-ов и пользовательские подписи не уникальны.
                if (frame == null ||
                    ErrorHandler.Failed(frame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out var documentPathObject)) ||
                    documentPathObject is not string documentPath ||
                    !string.Equals(documentPath, this.Document.FullName, StringComparison.OrdinalIgnoreCase)) {

                    continue;
                }

                // VSFPROPID_Hierarchy указывает не на «владельца» файла на диске, а на project
                // representation текущего editor frame. Для Ctrl+G по include это DxPlayer.
                if (ErrorHandler.Failed(frame.GetProperty((int)__VSFPROPID.VSFPROPID_Hierarchy, out var hierarchyObject)) ||
                    hierarchyObject is not IVsHierarchy hierarchy) {

                    continue;
                }

                var project = Utils.EnvDteUtils.GetDteProjectFromHierarchy(hierarchy);
                // Miscellaneous Files означает открытие файла вне solution и не является
                // полезным project context-ом для группировки.
                if (project != null && !Utils.EnvDteUtils.IsMiscProject(project)) {
                    return project.Name;
                }
            }

            return null;
        }

        private bool ProjectContainsDocumentInProject(EnvDTE.Project project) {
            ThreadHelper.ThrowIfNotOnUIThread();

            try {
                foreach (EnvDTE.ProjectItem item in project.ProjectItems) {
                    if (this.ProjectItemContainsDocument(item)) {
                        return true;
                    }
                }
            }
            catch {
                // Игнорируем ошибки проверки
            }

            return false;
        }


        // Метод проверки документа внутри ProjectItem (включая вложенные)
        private bool ProjectItemContainsDocument(EnvDTE.ProjectItem item) {
            ThreadHelper.ThrowIfNotOnUIThread();

            try {
                if (item.FileCount > 0) {
                    for (short i = 1; i <= item.FileCount; i++) {
                        string filePath = item.FileNames[i];
                        if (string.Equals(filePath, this.Document.FullName, StringComparison.OrdinalIgnoreCase)) {
                            return true;
                        }
                    }
                }

                // Проверяем вложенные элементы (вложенные папки, ссылки)
                if (item.ProjectItems?.Count > 0) {
                    foreach (EnvDTE.ProjectItem subItem in item.ProjectItems) {
                        if (this.ProjectItemContainsDocument(subItem)) {
                            return true;
                        }
                    }
                }
            }
            catch {
                // Игнорируем ошибки
            }

            return false;
        }
    }
}
