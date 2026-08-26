using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace TabsManagerExtension.VsShell.Solution.Services {
    /// <summary>
    /// Выполняет лёгкий лексический разбор C/C++ preprocessor directives без запуска компилятора.
    /// </summary>
    /// <remarks>
    /// Парсер намеренно не вычисляет макросы и условия препроцессора. Include внутри любой ветки
    /// \#if/\#ifdef/\#ifndef сохраняется с <see cref="Document.IncludeEntry.ConditionUnknown"/>.
    /// Это позволяет не потерять потенциальную зависимость, не притворяясь полноценным C++ frontend.
    /// </remarks>
    internal static class CppIncludeParser {
        public static List<Document.IncludeEntry> ParseFile(string filePath) {
            using var reader = new StreamReader(filePath);
            return Parse(reader);
        }


        internal static List<Document.IncludeEntry> Parse(TextReader reader) {
            var result = new List<Document.IncludeEntry>();
            var logicalLine = new StringBuilder();
            bool inBlockComment = false;
            int conditionDepth = 0;

            while (true) {
                string? physicalLine = reader.ReadLine();
                if (physicalLine == null) {
                    break;
                }

                string codeWithoutComments = RemoveComments(physicalLine, ref inBlockComment);
                logicalLine.Append(codeWithoutComments);

                // Обрабатываем простое line splicing: #include \\ + следующая строка.
                // Это не полноценная реализация translation phases C++, но покрывает обычную
                // запись многострочных preprocessor directives.
                if (RemoveTrailingLineContinuation(logicalLine)) {
                    continue;
                }

                ProcessLogicalLine(logicalLine.ToString(), ref conditionDepth, result);
                logicalLine.Clear();
            }

            if (logicalLine.Length > 0) {
                ProcessLogicalLine(logicalLine.ToString(), ref conditionDepth, result);
            }

            return result;
        }


        private static void ProcessLogicalLine(
            string line,
            ref int conditionDepth,
            List<Document.IncludeEntry> result) {

            if (!TryReadDirective(line, out string directive, out string operand)) {
                return;
            }

            if (directive.Equals("include", StringComparison.Ordinal)) {
                Document.IncludeEntry? includeEntry = ParseIncludeOperand(operand, conditionDepth > 0);
                if (includeEntry != null) {
                    result.Add(includeEntry);
                }
                return;
            }

            if (directive.Equals("if", StringComparison.Ordinal) ||
                directive.Equals("ifdef", StringComparison.Ordinal) ||
                directive.Equals("ifndef", StringComparison.Ordinal)) {

                conditionDepth++;
                return;
            }

            if (directive.Equals("endif", StringComparison.Ordinal) && conditionDepth > 0) {
                conditionDepth--;
            }
        }


        private static bool TryReadDirective(string line, out string directive, out string operand) {
            directive = string.Empty;
            operand = string.Empty;

            int cursor = 0;
            SkipWhitespace(line, ref cursor);
            if (cursor >= line.Length || line[cursor] != '#') {
                return false;
            }

            cursor++;
            SkipWhitespace(line, ref cursor);

            int directiveStart = cursor;
            while (cursor < line.Length && (char.IsLetter(line[cursor]) || line[cursor] == '_')) {
                cursor++;
            }

            if (cursor == directiveStart) {
                return false;
            }

            directive = line.Substring(directiveStart, cursor - directiveStart);
            operand = line.Substring(cursor).Trim();
            return true;
        }


        private static Document.IncludeEntry? ParseIncludeOperand(string operand, bool conditionUnknown) {
            if (string.IsNullOrWhiteSpace(operand)) {
                return null;
            }

            if (operand[0] == '"') {
                int closingQuote = operand.IndexOf('"', 1);
                return CreateDelimitedEntry(
                    operand,
                    closingQuote,
                    Document.IncludeKind.Quote,
                    conditionUnknown
                );
            }

            if (operand[0] == '<') {
                int closingAngle = operand.IndexOf('>', 1);
                return CreateDelimitedEntry(
                    operand,
                    closingAngle,
                    Document.IncludeKind.Angle,
                    conditionUnknown
                );
            }

            // Макрос может разворачиваться в "Header.h" или <Header.h>, но без набора define'ов
            // активной конфигурации угадать его значение нельзя. Сохраняем выражение явно,
            // а resolver пометит такую запись как MacroExpression, не выдавая случайный путь.
            return new Document.IncludeEntry(operand, Document.IncludeKind.Macro, conditionUnknown);
        }


        private static Document.IncludeEntry? CreateDelimitedEntry(
            string operand,
            int closingDelimiter,
            Document.IncludeKind kind,
            bool conditionUnknown) {

            if (closingDelimiter <= 1) {
                return null;
            }

            string rawInclude = operand.Substring(1, closingDelimiter - 1).Trim();
            if (string.IsNullOrWhiteSpace(rawInclude)) {
                return null;
            }

            return new Document.IncludeEntry(rawInclude, kind, conditionUnknown);
        }


        private static string RemoveComments(string line, ref bool inBlockComment) {
            var result = new StringBuilder(line.Length);
            bool inString = false;
            bool inCharacter = false;
            bool escaped = false;

            for (int index = 0; index < line.Length; index++) {
                char current = line[index];
                char next = index + 1 < line.Length ? line[index + 1] : '\0';

                if (inBlockComment) {
                    if (current == '*' && next == '/') {
                        inBlockComment = false;
                        index++;
                    }
                    continue;
                }

                if (inString || inCharacter) {
                    result.Append(current);
                    if (escaped) {
                        escaped = false;
                    }
                    else if (current == '\\') {
                        escaped = true;
                    }
                    else if (inString && current == '"') {
                        inString = false;
                    }
                    else if (inCharacter && current == '\'') {
                        inCharacter = false;
                    }
                    continue;
                }

                if (current == '/' && next == '/') {
                    break;
                }

                if (current == '/' && next == '*') {
                    // По правилам препроцессора комментарий заменяется пробелом. Это важно для
                    // #inc/**/lude: такая строка не должна склеиться в валидный #include.
                    result.Append(' ');
                    inBlockComment = true;
                    index++;
                    continue;
                }

                result.Append(current);
                if (current == '"') {
                    inString = true;
                }
                else if (current == '\'') {
                    inCharacter = true;
                }
            }

            return result.ToString();
        }


        private static bool RemoveTrailingLineContinuation(StringBuilder line) {
            int cursor = line.Length - 1;
            while (cursor >= 0 && char.IsWhiteSpace(line[cursor])) {
                cursor--;
            }

            if (cursor < 0 || line[cursor] != '\\') {
                return false;
            }

            line.Length = cursor;
            return true;
        }


        private static void SkipWhitespace(string value, ref int cursor) {
            while (cursor < value.Length && char.IsWhiteSpace(value[cursor])) {
                cursor++;
            }
        }
    }
}
