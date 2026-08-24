using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace TabsManagerExtension {
    namespace Enums {
        public enum AnchorKind {
            Section,
            Subsection,
            Separator,
        }
    }
}

namespace TabsManagerExtension.State.TextEditor {
    public class AnchorPoint {
        public string Title { get; set; }
        public int LineNumber { get; set; }
        public Enums.AnchorKind AnchorKind { get; set; }

        public AnchorPoint(string title, int lineNumber, Enums.AnchorKind kind = Enums.AnchorKind.Subsection) {
            this.Title = title;
            this.LineNumber = lineNumber;
            this.AnchorKind = kind;
        }
    }


    public static class AnchorParser {
        // Для каждой позиции строится короткое окно строк: секционный шаблон может захватывать заголовок
        // вместе со следующей строкой-разделителем, а шаблон подпункта — только текущую строку.
        public class SourceLine {
            public int Index { get; }
            public string Text { get; }

            public SourceLine(int index, string text) {
                this.Index = index;
                this.Text = text;
            }
        }

        public static List<AnchorPoint> ParseLinesWithContextWindow(
            List<string> lines,
            string sectionPattern,
            string subsectionPattern
        ) {
            var result = new List<AnchorPoint>();
            var sectionRegex = CreateRegex(sectionPattern);
            var subsectionRegex = CreateRegex(subsectionPattern);
            int i = 0;

            while (i < lines.Count) {
                var context = BuildContext(lines, i, linesAfter: 3);
                var anchor = TryParseAnchor(context, sectionRegex, subsectionRegex, out int linesConsumed);

                if (anchor != null) {
                    result.Add(anchor);
                    i += linesConsumed;
                }
                else {
                    i++;
                }
            }

            return result;
        }

        public static bool ContainsAnchor(
            List<string> lines,
            string sectionPattern,
            string subsectionPattern
        ) {
            var sectionRegex = CreateRegex(sectionPattern);
            var subsectionRegex = CreateRegex(subsectionPattern);

            for (int i = 0; i < lines.Count; i++) {
                var context = BuildContext(lines, i, linesAfter: 3);
                if (TryParseAnchor(context, sectionRegex, subsectionRegex, out _) != null) {
                    return true;
                }
            }

            return false;
        }

        private static List<SourceLine> BuildContext(List<string> lines, int startIndex, int linesAfter) {
            var context = new List<SourceLine>();

            int end = Math.Min(lines.Count - 1, startIndex + linesAfter);
            for (int i = startIndex; i <= end; i++) {
                context.Add(new SourceLine(i, lines[i]));
            }

            return context;
        }

        private static Regex CreateRegex(string pattern) {
            // Ограничение времени защищает UI от слишком дорогих пользовательских регулярных выражений.
            return new Regex(pattern, RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
        }

        private static AnchorPoint? TryParseAnchor(
            List<SourceLine> contextLines,
            Regex sectionRegex,
            Regex subsectionRegex,
            out int linesConsumed
        ) {
            linesConsumed = 1;

            if (contextLines.Count == 0) {
                return null;
            }

            var first = contextLines[0];
            // Шаблоны применяются с начала объединённого окна, поэтому номер якоря соответствует first.Index.
            string contextText = string.Join(Environment.NewLine, contextLines.Select(sourceLine => sourceLine.Text));

            var sectionMatch = TryMatch(sectionRegex, contextText);
            if (TryGetTitle(sectionMatch, out string sectionTitle)) {
                linesConsumed = CountConsumedLines(sectionMatch.Value);
                return new AnchorPoint(sectionTitle, first.Index + 1, Enums.AnchorKind.Section);
            }

            var subsectionMatch = TryMatch(subsectionRegex, contextText);
            if (TryGetTitle(subsectionMatch, out string subsectionTitle)) {
                linesConsumed = CountConsumedLines(subsectionMatch.Value);
                return new AnchorPoint(subsectionTitle, first.Index + 1, Enums.AnchorKind.Subsection);
            }

            return null;
        }

        private static Match TryMatch(Regex regex, string text) {
            try {
                return regex.Match(text);
            }
            catch (RegexMatchTimeoutException) {
                // Пользовательский шаблон не должен блокировать UI редактора.
                return Match.Empty;
            }
        }

        private static bool TryGetTitle(Match match, out string title) {
            title = match.Success && match.Index == 0 ? match.Groups["title"].Value.Trim() : string.Empty;
            return !string.IsNullOrEmpty(title);
        }

        private static int CountConsumedLines(string matchedText) {
            return Math.Max(1, matchedText.Count(character => character == '\n') + 1);
        }

        public static List<AnchorPoint> InsertSeparators(List<AnchorPoint> anchors) {
            var result = new List<AnchorPoint>();

            for (int i = 0; i < anchors.Count; i++) {
                var current = anchors[i];
                result.Add(current);

                bool isLast = i == anchors.Count - 1;
                bool nextIsSection = !isLast && anchors[i + 1].AnchorKind == Enums.AnchorKind.Section;

                if (current.AnchorKind == Enums.AnchorKind.Subsection && nextIsSection) {
                    result.Add(new AnchorPoint(string.Empty, -1, Enums.AnchorKind.Separator));
                }
            }

            return result;
        }
    }
}
