using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;
using Microsoft.VisualStudio.Text.Formatting;
using Microsoft.VisualStudio.Text.Projection; // BufferGraph.MapUpToSnapshot

namespace TabsManagerExtension.VsShell.TextEditor {
    /// <summary>
    /// Объявляем собственный слой адорнеров, который будет отрисовывать формулы поверх текста.
    /// </summary>
    internal static class FormulaAdornmentLayerDefinition {
        public const string LayerName = "VsixFormulaAdornerLayer";

        [Export(typeof(AdornmentLayerDefinition))]
        [Name(LayerName)]
        [Order(After = PredefinedAdornmentLayers.Text, Before = PredefinedAdornmentLayers.Caret)]
        [TextViewRole(PredefinedTextViewRoles.Document)]
        public static AdornmentLayerDefinition? Definition;
    }

    /// <summary>
    /// Подключаем менеджер при создании нового IWpfTextView (окна редактора).
    /// </summary>
    [Export(typeof(IWpfTextViewCreationListener))]
    [ContentType("text")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    internal sealed class FormulaAdornmentTextViewCreationListener : IWpfTextViewCreationListener {
        public void TextViewCreated(IWpfTextView textView) {
            // Храним один менеджер на вью — пока живо вью, жив менеджер.
            textView.Properties.GetOrCreateSingletonProperty(
                typeof(FormulaAdornmentManager),
                () => new FormulaAdornmentManager(textView)
            );
        }
    }

    /// <summary>
    /// Менеджер формул.
    /// Основные этапы пайплайна:
    ///  1) На каждый Change/Layout разбиваем документ на «вертикальные блоки» — подряд идущие строки,
    ///     где встречается хотя бы один символ вертикали '|'. Табы и пробелы учитываются одинаково
    ///     (позиции считаются в «визуальных колонках» относительно TabSize).
    ///  2) Внутри каждого вертикального блока берём **унион** всех визуальных колонок '|'.
    ///     Из отсортированного списка колонок строим **неперекрывающиеся пары** (0–1, 2–3, ...).
    ///  3) Для каждой пары колонок ищем все вертикальные «пробеги» (поддиапазоны строк),
    ///     где обе колонки присутствуют подряд на каждой строке. Минимум 3 строки.
    ///  4) Для каждого пробега строим ViewBox:
    ///       - границы по реальным bounds символов '|' (без самих линий, с небольшим внутренним отступом);
    ///       - спан привязки берём между вертикалями на **верхней** строке пробега, мэпим в snapshot вью
    ///         через BufferGraph.MapUpToSnapshot (проекции, diff и т.п.).
    ///  5) В каждый ViewBox кладём формулу, **вписывая её равномерно (Uniform)** как в видеоплеерах
    ///     («letterbox»): без искажения aspect ratio и без апскейла (не увеличиваем формулу сверх натурального).
    /// </summary>
    internal sealed class FormulaAdornmentManager : IDisposable {
        private readonly IWpfTextView _view;
        private readonly IAdornmentLayer _layer;
        private readonly object _tag = new object();

        // Прямоугольник в координатах View + SnapshotSpan для текстовой привязки.
        private readonly struct ViewBox {
            public readonly SnapshotSpan ViewSpan;
            public readonly double Left;
            public readonly double Top;
            public readonly double Width;
            public readonly double Height;
            public ViewBox(SnapshotSpan viewSpan, double left, double top, double width, double height) {
                ViewSpan = viewSpan;
                Left = left;
                Top = top;
                Width = width;
                Height = height;
            }
        }

        // Позиция вертикальной черты на строке: индекс в тексте и «визуальная колонка» с учётом табов.
        private readonly struct PipePos {
            public readonly int Index;
            public readonly int Col;
            public PipePos(int index, int col) {
                Index = index;
                Col = col;
            }
        }

        public FormulaAdornmentManager(IWpfTextView view) {
            _view = view ?? throw new ArgumentNullException(nameof(view));
            _layer = _view.GetAdornmentLayer(FormulaAdornmentLayerDefinition.LayerName)
                     ?? throw new InvalidOperationException("Adornment layer not found");

            // Подписки:
            //  - LayoutChanged  → пересчитываем позицию формул при скролле/масштабе/сворачивании регионов;
            //  - Changed (DocumentBuffer) → реагируем на изменение текста.
            _view.Closed += this.OnClosed;
            _view.LayoutChanged += this.OnLayoutChanged;
            _view.TextDataModel.DocumentBuffer.Changed += this.OnBufferChanged;

            this.EnsureOverlay();
        }

        public void Dispose() {
            _view.TextDataModel.DocumentBuffer.Changed -= this.OnBufferChanged;
            _view.LayoutChanged -= this.OnLayoutChanged;
            _view.Closed -= this.OnClosed;
            _layer.RemoveAdornmentsByTag(_tag);
        }

        private void OnClosed(object? s, EventArgs e) {
            this.Dispose();
        }

        private void OnLayoutChanged(object? s, TextViewLayoutChangedEventArgs e) {
            // Здесь у вью уже готов актуальный набор TextViewLines для собственного snapshot.
            this.EnsureOverlay();
        }

        private void OnBufferChanged(object? s, TextContentChangedEventArgs e) {
            // На каждую правку просто перерисовываем. Если строки ещё не готовы —
            // часть ViewBox'ов пропустим, и следующий LayoutChanged их дорисует.
            this.EnsureOverlay();
        }

        /// <summary>
        /// Удаляем старые адорнеры и добавляем новые по актуальному набору ViewBox'ов.
        /// </summary>
        private void EnsureOverlay() {
            _layer.RemoveAdornmentsByTag(_tag);

            var boxes = this.FindPipeBoxesInView();
            if (boxes.Count == 0) {
                return;
            }

            for (int k = 0; k < boxes.Count; ++k) {
                var element = this.CreateCenteredFormulaElement(boxes[k]);
                _layer.AddAdornment(
                    AdornmentPositioningBehavior.TextRelative,
                    boxes[k].ViewSpan, // текстовая привязка — верхняя строка текущей «ячейки»
                    _tag,
                    element,
                    removedCallback: null
                );
            }
        }

        /// <summary>
        /// Создаёт визуальный элемент формулы и центрирует его **равномерно** («letterbox») внутри ViewBox.
        /// Масштабируем только вниз (без апскейла), aspect ratio сохраняется.
        /// </summary>
        private UIElement CreateCenteredFormulaElement(ViewBox vb) {
            var element = new WpfMath.Controls.FormulaControl {
                Formula = @"\frac{a + b}{x}",
                FontSize = 22,
                Foreground = Brushes.Orange,
                Background = Brushes.Transparent,
                Padding = new Thickness(2),
                IsHitTestVisible = false,
                SnapsToDevicePixels = true
            };

            // Натуральный размер при текущем FontSize
            element.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            var natural = element.DesiredSize;
            if (natural.Width <= 0.0 || natural.Height <= 0.0) {
                return element;
            }

            // Внутренний отступ от вертикалей
            const double innerPad = 2.0;
            var boxW = Math.Max(0.0, vb.Width - innerPad * 2.0);
            var boxH = Math.Max(0.0, vb.Height - innerPad * 2.0);

            // Uniform-вписывание без апскейла
            var scale = Math.Min(1.0, Math.Min(boxW / natural.Width, boxH / natural.Height));
            if (double.IsNaN(scale) || double.IsInfinity(scale) || scale <= 0.0) {
                scale = 1.0;
            }

            var scaledW = natural.Width * scale;
            var scaledH = natural.Height * scale;

            // Центрирование
            var left = vb.Left + (vb.Width - scaledW) * 0.5;
            var top = vb.Top + (vb.Height - scaledH) * 0.5;

            var tg = new TransformGroup();
            tg.Children.Add(new ScaleTransform(scale, scale));  // сначала масштаб
            tg.Children.Add(new TranslateTransform(left, top)); // затем перенос
            element.RenderTransform = tg;

            return element;
        }

        // === ХЕЛПЕРЫ для вычисления колонок и индексов '|' ========================================

        // Возвращает позиции всех '|' на строке в «визуальных колонках» (табы разворачиваются до TabSize).
        private List<PipePos> GetPipePositions(string text, int tabSize) {
            var list = new List<PipePos>(Math.Max(2, text.Length / 8));
            int col = 0;
            for (int i = 0; i < text.Length; ++i) {
                char c = text[i];
                if (c == '\t') {
                    int step = tabSize - (col % tabSize);
                    col += step; // прыжок до следующего таб-стопа
                }
                else {
                    if (c == '|') {
                        list.Add(new PipePos(i, col));
                    }
                    col += 1;
                }
            }
            return list;
        }

        private static int FindIndexByCol(List<PipePos> list, int col) {
            for (int k = 0; k < list.Count; ++k) {
                if (list[k].Col == col) {
                    return list[k].Index;
                }
            }
            return -1;
        }

        private static int FindLastIndexByCol(List<PipePos> list, int col) {
            for (int k = list.Count - 1; k >= 0; --k) {
                if (list[k].Col == col) {
                    return list[k].Index;
                }
            }
            return -1;
        }

        // === ДЕТЕКТОР ВСЕХ «ЯЧЕЕК» МЕЖДУ ВЕРТИКАЛЯМИ =============================================

        /// <summary>
        /// Находит все ViewBox'ы по всему документу.
        /// Алгоритм устойчив к табам/пробелам (сравнение по визуальным колонкам) и к проекциям (MapUpToSnapshot).
        /// В одном вертикальном блоке формируем **неперекрывающиеся пары** колонок: (0–1), (2–3), (4–5), ...
        /// Для каждой пары строим все вертикальные «пробеги» (минимум 3 строки подряд), где обе колонки присутствуют.
        /// </summary>
        private List<ViewBox> FindPipeBoxesInView() {
            var result = new List<ViewBox>();

            var docSnap = _view.TextDataModel.DocumentBuffer.CurrentSnapshot;
            if (docSnap.Length == 0) {
                return result;
            }

            int tabSize = _view.Options.GetOptionValue(DefaultOptions.TabSizeOptionId);
            if (tabSize <= 0) {
                tabSize = 4;
            }

            int i = 0;
            while (i < docSnap.LineCount) {
                // 1) Ищем старт вертикального блока: первая строка, где есть хотя бы один '|'
                var firstLineInScan = docSnap.GetLineFromLineNumber(i);
                var firstPipes = this.GetPipePositions(firstLineInScan.GetText(), tabSize);
                if (firstPipes.Count == 0) {
                    i++;
                    continue;
                }

                // 2) Собираем подряд строки, в которых встречается хотя бы один '|'
                int start = i;
                int end = i;

                var colsPerLine = new List<Dictionary<int, int>>(); // на каждой строке: col → index
                var allColsSet = new HashSet<int>();

                int j = i;
                for (; j < docSnap.LineCount; ++j) {
                    var l = docSnap.GetLineFromLineNumber(j);
                    var pp = this.GetPipePositions(l.GetText(), tabSize);
                    if (pp.Count == 0) {
                        break; // блок закончился перед j
                    }

                    var dict = new Dictionary<int, int>();
                    for (int q = 0; q < pp.Count; ++q) {
                        int col = pp[q].Col;
                        if (!dict.ContainsKey(col)) {
                            dict[col] = pp[q].Index; // сохраняем левый индекс для колонки
                        }
                    }

                    colsPerLine.Add(dict);
                    foreach (var col in dict.Keys) {
                        allColsSet.Add(col);
                    }

                    end = j;
                }

                int linesInBlock = end - start + 1;
                if (linesInBlock >= 3 && allColsSet.Count >= 2) {
                    // 3) Отсортированный унион всех колонок '|' в блоке
                    var allCols = new List<int>(allColsSet);
                    allCols.Sort();

                    // 4) Неперекрывающиеся пары колонок: (0–1), (2–3), (4–5), ...
                    for (int c = 0; c + 1 < allCols.Count; c += 2) {
                        int leftCol = allCols[c];
                        int rightCol = allCols[c + 1];

                        // 5) Все вертикальные пробеги по строкам, где обе колонки присутствуют подряд
                        int runStart = -1;

                        for (int row = 0; row < colsPerLine.Count; ++row) {
                            var dict = colsPerLine[row];
                            bool hasLeft = dict.ContainsKey(leftCol);
                            bool hasRight = dict.ContainsKey(rightCol);

                            if (hasLeft && hasRight) {
                                if (runStart < 0) {
                                    runStart = row;
                                }
                            }
                            else {
                                // Закрываем пробег [runStart..row-1]
                                if (runStart >= 0) {
                                    int runEnd = row - 1;
                                    this.TryAddViewBoxForRun(result, docSnap, start, runStart, runEnd, leftCol, rightCol, tabSize);
                                    runStart = -1;
                                }
                            }
                        }

                        // Хвост пробега в конце блока
                        if (runStart >= 0) {
                            this.TryAddViewBoxForRun(result, docSnap, start, runStart, colsPerLine.Count - 1, leftCol, rightCol, tabSize);
                        }
                    }
                }

                // 6) Переходим к строке после текущего вертикального блока
                i = Math.Max(end + 1, i + 1);
            }

            return result;
        }

        /// <summary>
        /// Собирает ViewBox для одного вертикального пробега строк между двумя колонками.
        /// Делает все проверки готовности (MapUpToSnapshot, наличие TextViewLines и т.д.).
        /// </summary>
        private void TryAddViewBoxForRun(
            List<ViewBox> sink,
            ITextSnapshot docSnap,
            int blockStartLine,
            int runStartRow,
            int runEndRow,
            int leftCol,
            int rightCol,
            int tabSize
        ) {
            if ((runEndRow - runStartRow + 1) < 3) {
                return; // слишком низкий пробег — игнорируем
            }

            var firstLine = docSnap.GetLineFromLineNumber(blockStartLine + runStartRow);
            var lastLine = docSnap.GetLineFromLineNumber(blockStartLine + runEndRow);

            // Индексы '|' на верхней/нижней строках для нужных визуальных колонок
            int topLeftIdx = FindIndexByCol(this.GetPipePositions(firstLine.GetText(), tabSize), leftCol);
            int topRightIdx = FindLastIndexByCol(this.GetPipePositions(firstLine.GetText(), tabSize), rightCol);
            int botLeftIdx = FindIndexByCol(this.GetPipePositions(lastLine.GetText(), tabSize), leftCol);
            int botRightIdx = FindLastIndexByCol(this.GetPipePositions(lastLine.GetText(), tabSize), rightCol);

            if (topLeftIdx < 0 || topRightIdx < 0 || botLeftIdx < 0 || botRightIdx < 0) {
                return;
            }

            // ViewSpan: участок между вертикалями на верхней строке пробега (в документе)
            var docStartPt = new SnapshotPoint(docSnap, firstLine.Start.Position + topLeftIdx + 1);
            var docEndPt = new SnapshotPoint(docSnap, firstLine.Start.Position + topRightIdx);
            var docSpan = new SnapshotSpan(docStartPt, docEndPt);

            // Мэпим в snapshot вью (проекции/diff/peek)
            var viewSnap = _view.TextSnapshot;
            var mapped = _view.BufferGraph.MapUpToSnapshot(docSpan, SpanTrackingMode.EdgeExclusive, viewSnap);
            if (mapped.Count == 0 || _view.TextViewLines == null || _view.TextViewLines.Count == 0) {
                return;
            }

            var viewSpan = mapped[0];

            // Верхняя и нижняя визуальные строки (для точных графических границ)
            ITextViewLine topLine;
            try {
                topLine = _view.GetTextViewLineContainingBufferPosition(viewSpan.Start);
            }
            catch {
                return;
            }

            var topLeftDoc = new SnapshotPoint(docSnap, firstLine.Start.Position + topLeftIdx);
            var topRightDoc = new SnapshotPoint(docSnap, firstLine.Start.Position + topRightIdx);
            var bottomLeftDoc = new SnapshotPoint(docSnap, lastLine.Start.Position + botLeftIdx);
            var bottomRightDoc = new SnapshotPoint(docSnap, lastLine.Start.Position + botRightIdx);

            var tl = _view.BufferGraph.MapUpToSnapshot(new SnapshotSpan(topLeftDoc, 0), SpanTrackingMode.EdgeExclusive, viewSnap);
            var tr = _view.BufferGraph.MapUpToSnapshot(new SnapshotSpan(topRightDoc, 0), SpanTrackingMode.EdgeExclusive, viewSnap);
            var bl = _view.BufferGraph.MapUpToSnapshot(new SnapshotSpan(bottomLeftDoc, 0), SpanTrackingMode.EdgeExclusive, viewSnap);
            var br = _view.BufferGraph.MapUpToSnapshot(new SnapshotSpan(bottomRightDoc, 0), SpanTrackingMode.EdgeExclusive, viewSnap);

            if (tl.Count == 0 || tr.Count == 0 || bl.Count == 0 || br.Count == 0) {
                return;
            }

            ITextViewLine bottomLine;
            try {
                bottomLine = _view.GetTextViewLineContainingBufferPosition(bl[0].Start);
            }
            catch {
                return;
            }

            // Реальные границы символов '|' сверху и снизу
            var tlB = topLine.GetCharacterBounds(tl[0].Start);
            var trB = topLine.GetCharacterBounds(tr[0].Start);
            var blB = bottomLine.GetCharacterBounds(bl[0].Start);
            var brB = bottomLine.GetCharacterBounds(br[0].Start);

            // Внутренние границы коробки (без самих вертикалей + небольшой отступ)
            const double innerPad = 2.0;
            var left = Math.Max(tlB.Right, blB.Right) + innerPad;
            var right = Math.Min(trB.Left, brB.Left) - innerPad;
            var top = Math.Max(tlB.Top, trB.Top) + innerPad;
            var bottom = Math.Min(blB.Bottom, brB.Bottom) - innerPad;

            if (right <= left || bottom <= top) {
                return;
            }

            sink.Add(new ViewBox(viewSpan, left, top, right - left, bottom - top));
        }
    }
}
