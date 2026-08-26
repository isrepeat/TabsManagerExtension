using System.IO;

namespace TabsManagerExtension.VsShell.Solution.Services {
    public static class IncludeResolverService {
        /// <summary>
        /// Пробует разрешить include-запись в абсолютный путь с учётом её delimiter.
        /// Возвращает путь к файлу, если он существует; иначе — null.
        /// </summary>
        public static string? TryResolveInclude(
            Document.IncludeEntry includeEntry,
            string includingFilePath,
            VsShell.Project.LoadedProject ownerProject,
            MsBuildSolutionWatcher msBuildSolutionWatcher) {

            if (includeEntry.Kind == Document.IncludeKind.Macro) {
                return null;
            }

            try {
                string includePath = includeEntry.RawInclude.Replace('/', '\\');

                // Quote include повторяет поведение MSVC: сначала каталог включающего файла,
                // затем include environment. Для angle include локальный каталог пропускается.
                // Например, <vector> не должен случайно разрешиться в Source/vector рядом с .cpp.
                if (includeEntry.Kind == Document.IncludeKind.Quote) {
                    string baseDir = Path.GetDirectoryName(includingFilePath)!;
                    string resolvedLocal = Path.GetFullPath(Path.Combine(baseDir, includePath));
                    if (File.Exists(resolvedLocal)) {
                        return resolvedLocal;
                    }
                }

                // Для обоих delimited-вариантов продолжаем поиск в evaluated include directories
                // активной Configuration|Platform проекта.
                var projectIncludeDirs = msBuildSolutionWatcher.GetCachedIncludeDirectoriesFor(ownerProject.FullName);
                foreach (var dir in projectIncludeDirs) {
                    string resolved = Path.GetFullPath(Path.Combine(dir, includePath));
                    if (File.Exists(resolved)) {
                        return resolved;
                    }
                }
            }
            catch {
                // ignore
            }

            return null;
        }


        /// <summary>
        /// Совместимый со старым API overload. Строка без delimiter трактуется как quote include.
        /// Для нового кода следует передавать <see cref="Document.IncludeEntry"/>, иначе различить
        /// <c>"Header.h"</c> и <c>&lt;Header.h&gt;</c> невозможно.
        /// </summary>
        public static string? TryResolveInclude(
            string includeRaw,
            string includingFilePath,
            VsShell.Project.LoadedProject ownerProject,
            MsBuildSolutionWatcher msBuildSolutionWatcher) {

            return TryResolveInclude(
                new Document.IncludeEntry(includeRaw),
                includingFilePath,
                ownerProject,
                msBuildSolutionWatcher
            );
        }

        ///// <summary>
        ///// Проверяет, разрешается ли includeRaw из includingFilePath в конкретный файл candidateFilePath.
        ///// </summary>
        //public static bool IncludeResolvesToFile(
        //    string includeRaw,
        //    string includingFilePath,
        //    string candidateFilePath,
        //    VsShell.Project.ShellProject ownerProject,
        //    MsBuildSolutionWatcher msBuildSolutionWatcher) {

        //    var resolved = TryResolveInclude(includeEntry, includingFilePath, ownerProject, msBuildSolutionWatcher);
        //    if (resolved == null) {
        //        return false;
        //    }

        //    return string.Equals(Path.GetFullPath(candidateFilePath), resolved, StringComparison.OrdinalIgnoreCase);
        //}
    }
}
