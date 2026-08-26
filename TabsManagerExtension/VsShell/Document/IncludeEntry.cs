using System;
using System.IO;

namespace TabsManagerExtension.VsShell.Document {
    public enum IncludeKind {
        Quote,
        Angle,
        Macro,
    }


    public enum IncludeResolutionFailureReason {
        NotFound,
        MacroExpression,
    }


    public class IncludeEntry {
        public string RawInclude { get; }
        public string NormalizedName { get; }
        public IncludeKind Kind { get; }
        public bool ConditionUnknown { get; }

        /// <summary>
        /// Создаёт quote include. Overload сохранён для совместимости со старым API.
        /// Новый parser всегда передаёт фактически найденный <see cref="IncludeKind"/> явно.
        /// </summary>
        public IncludeEntry(string rawInclude)
            : this(rawInclude, IncludeKind.Quote, false) { }

        public IncludeEntry(string rawInclude, IncludeKind kind)
            : this(rawInclude, kind, false) { }

        public IncludeEntry(string rawInclude, IncludeKind kind, bool conditionUnknown) {
            this.RawInclude = rawInclude;
            this.NormalizedName = Path.GetFileName(rawInclude);
            this.Kind = kind;
            this.ConditionUnknown = conditionUnknown;
        }

        public override bool Equals(object? obj) {
            return obj is IncludeEntry other &&
                   StringComparer.OrdinalIgnoreCase.Equals(this.RawInclude, other.RawInclude) &&
                   this.Kind == other.Kind &&
                   this.ConditionUnknown == other.ConditionUnknown;
        }

        public override int GetHashCode() {
            int result = StringComparer.OrdinalIgnoreCase.GetHashCode(this.RawInclude);
            result = (result * 397) ^ this.Kind.GetHashCode();
            result = (result * 397) ^ this.ConditionUnknown.GetHashCode();
            return result;
        }

        public override string ToString() {
            return this.Kind switch {
                IncludeKind.Quote => $"\"{this.RawInclude}\"",
                IncludeKind.Angle => $"<{this.RawInclude}>",
                _ => this.RawInclude,
            };
        }
    }


    public class ResolvedIncludeEntry {
        public IncludeEntry IncludeEntry { get; }
        public string? ResolvedPath { get; }
        public IncludeResolutionFailureReason? FailureReason { get; }

        public ResolvedIncludeEntry(IncludeEntry includeEntry, string? resolvedPath) {
            this.IncludeEntry = includeEntry;
            this.ResolvedPath = resolvedPath;
            this.FailureReason = resolvedPath != null
                ? null
                : includeEntry.Kind == IncludeKind.Macro
                    ? IncludeResolutionFailureReason.MacroExpression
                    : IncludeResolutionFailureReason.NotFound;
        }

        public override bool Equals(object? obj) {
            return obj is ResolvedIncludeEntry other &&
                   this.IncludeEntry.Equals(other.IncludeEntry) &&
                   StringComparer.OrdinalIgnoreCase.Equals(this.ResolvedPath, other.ResolvedPath);
        }

        public override int GetHashCode() {
            int h1 = this.IncludeEntry.GetHashCode();
            int h2 = this.ResolvedPath is null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(this.ResolvedPath);
            return (h1 * 397) ^ h2;
        }

        public override string ToString() {
            string resolution = this.ResolvedPath ?? $"unresolved ({this.FailureReason})";
            return $"{this.IncludeEntry} → {resolution}";
        }
    }
}
