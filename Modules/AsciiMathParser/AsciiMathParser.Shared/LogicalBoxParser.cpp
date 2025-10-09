#include "LogicalBoxParser.h"
#include <algorithm>

namespace AsciiMathParser {
	namespace Core {
		static int32_t VisualColOf(const std::u16string& s, int32_t tabSize, int32_t charIdx) {
			int32_t col = 0;
			for (int32_t i = 0; i < charIdx && i < static_cast<int32_t>(s.size()); i++) {
				char16_t ch = s[static_cast<size_t>(i)];
				if (ch == u'\t') {
					int32_t next = ((col / tabSize) + 1) * tabSize;
					col = next;
				}
				else {
					col++;
				}
			}
			return col;
		}

		void FindLogicalBoxes(
			const std::vector<LineU16>& lines,
			int32_t tabSize,
			std::vector<LogicalBox>& outBoxes
		) {
			outBoxes.clear();
			if (lines.size() < 3) {
				return;
			}

			// ЗАГЛУШКА: найдём пары вертикалей '|' на каждой строке и
			// если одна и та же пара колонок повторяется >= 3 строк подряд — вернём коробку.
			// Реальную вашу логику (унион колонок, пробеги, вложенности) сюда.

			struct Pair { int32_t lCol, rCol, topLine, bottomLine; int32_t tlIdx, trIdx, blIdx, brIdx; };

			std::vector<Pair> run;
			run.reserve(lines.size());

			auto flushRun = [&]() {
				if (run.size() >= 3) {
					Pair p = run.front();
					Pair q = run.back();
					LogicalBox box{};
					box.topLine = p.topLine;
					box.bottomLine = q.bottomLine;
					box.leftCol = p.lCol;
					box.rightCol = p.rCol;
					box.topLeftCharIdx = p.tlIdx;
					box.topRightCharIdx = p.trIdx;
					box.bottomLeftCharIdx = q.blIdx;
					box.bottomRightCharIdx = q.brIdx;
					outBoxes.push_back(box);
				}
				run.clear();
				};

			for (int32_t i = 0; i < static_cast<int32_t>(lines.size()); i++) {
				const auto& s = lines[static_cast<size_t>(i)].text;
				std::vector<int32_t> barsIdx;
				barsIdx.reserve(16);
				for (int32_t j = 0; j < static_cast<int32_t>(s.size()); j++) {
					if (s[static_cast<size_t>(j)] == u'|') {
						barsIdx.push_back(j);
					}
				}
				if (barsIdx.size() < 2) {
					flushRun();
					continue;
				}

				// Возьмём первую непересекающуюся пару (0,1) как пример
				int32_t tl = barsIdx[0];
				int32_t tr = barsIdx[1];
				int32_t lCol = VisualColOf(s, tabSize, tl);
				int32_t rCol = VisualColOf(s, tabSize, tr);

				if (!run.empty()) {
					if (run.back().lCol == lCol && run.back().rCol == rCol) {
						// Продолжаем пробег
						run.back().bottomLine = i;
						run.back().blIdx = tl;
						run.back().brIdx = tr;
					}
					else {
						flushRun();
						Pair p{ lCol, rCol, i, i, tl, tr, tl, tr };
						run.push_back(p);
					}
				}
				else {
					Pair p{ lCol, rCol, i, i, tl, tr, tl, tr };
					run.push_back(p);
				}
			}
			flushRun();
		}
	}
}