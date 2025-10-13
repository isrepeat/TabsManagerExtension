#pragma once
#include "../Model/Geometry.h"
#include "../Model/Grid.h"
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		namespace Detect {
			struct FractionBar {
				int y;
				int x1;
				int x2;
				Model::Region numRegion;
				Model::Region denRegion;
			};

			class FractionBarDetector {
			public:
				std::vector<FractionBar> DetectBars(const Model::AsciiGrid& grid) const {
					std::vector<FractionBar> bars;

					const int rows = grid.Height();
					const int cols = grid.Width();

					// ============================
					// ФАЗА 1: собрать все черты
					// ============================
					struct RawBar {
						int y;
						int x1;
						int x2;
					};

					std::vector<RawBar> rawBars;
					rawBars.reserve(static_cast<std::size_t>(rows));

					for (int y = 0; y < rows; ++y) {
						int x = 0;

						while (x < cols) {
							// найти начало сегмента черты
							while (x < cols && !this->IsHorizontalBarChar(grid.At(y, x))) {
								++x;
							}

							if (x >= cols) {
								break;
							}

							const int startX = x;

							while (x < cols && this->IsHorizontalBarChar(grid.At(y, x))) {
								++x;
							}

							const int endX = x - 1;

							// верификация: сверху и снизу есть непустые символы в коридоре [startX..endX]
							bool hasAbove = false;
							bool hasBelow = false;

							if (y > 0) {
								const bool allSpace = Model::Geometry::IsRowRangeAllSpace(
									grid,
									y - 1,
									startX,
									endX
								);

								if (!allSpace) {
									hasAbove = true;
								}
							}

							if (y + 1 < rows) {
								const bool allSpace = Model::Geometry::IsRowRangeAllSpace(
									grid,
									y + 1,
									startX,
									endX
								);

								if (!allSpace) {
									hasBelow = true;
								}
							}

							if (hasAbove && hasBelow) {
								rawBars.push_back(RawBar{
									.y = y,
									.x1 = startX,
									.x2 = endX
									});
							}
						}
					}

					if (rawBars.empty()) {
						return bars;
					}

					// Быстрый доступ: на каждой строке держим слитые интервалы черт.
					std::vector<std::vector<Model::SpanX>> barSpansByRow(
						static_cast<std::size_t>(rows)
					);

					for (const auto& b : rawBars) {
						barSpansByRow[static_cast<std::size_t>(b.y)]
							.push_back(Model::SpanX{ b.x1, b.x2 });
					}

					for (auto& spans : barSpansByRow) {
						std::sort(
							spans.begin(),
							spans.end(),
							[](const Model::SpanX& a, const Model::SpanX& b) {
								if (a.x1 != b.x1) {
									return a.x1 < b.x1;
								}
								else {
									return a.x2 < b.x2;
								}
							}
						);

						std::vector<Model::SpanX> merged;

						for (const auto& s : spans) {
							if (merged.empty()) {
								merged.push_back(s);
							}
							else {
								auto& last = merged.back();

								if (s.x1 <= last.x2 + 1) {
									if (s.x2 > last.x2) {
										last.x2 = s.x2;
									}
								}
								else {
									merged.push_back(s);
								}
							}
						}

						spans = std::move(merged);
					}

					auto row_has_overlapping_bar = [&](int y, int x1, int x2) -> bool {
						if (y < 0 || y >= rows) {
							return false;
						}

						const auto& spans = barSpansByRow[static_cast<std::size_t>(y)];

						for (const auto& s : spans) {
							const int L = std::max(s.x1, x1);
							const int R = std::min(s.x2, x2);

							if (L <= R) {
								return true;
							}
						}

						return false;
						};

					// ==========================================
					// ФАЗА 2: посчитать окна Num/Den для каждого
					// ==========================================
					bars.reserve(rawBars.size());

					for (const auto& b : rawBars) {
						const int y = b.y;
						const int x1 = b.x1;
						const int x2 = b.x2;

						// --- верхняя граница (Num):
						//     поднимаемся вверх; стоп на ПУСТОЙ строке
						//     или на ЧЕРТЕ, пересекающей [x1..x2].
						//     ВАЖНО: если встретили ЧЕРТУ — оставляем границу НА этой строке,
						//     чтобы вложенная черта попала в регион и распарсилась как Frac.
						int numTop = y - 1;

						while (numTop >= 0) {
							if (row_has_overlapping_bar(numTop, x1, x2)) {
								break;
							}

							const bool allSpace = Model::Geometry::IsRowRangeAllSpace(
								grid,
								numTop,
								x1,
								x2
							);

							if (allSpace) {
								++numTop;
								break;
							}

							--numTop;
						}

						if (numTop < 0) {
							numTop = 0;
						}

						const int numBottom = y - 1;

						// --- нижняя граница (Den):
						//     спускаемся вниз; стоп на ПУСТОЙ строке
						//     или на ЧЕРТЕ, пересекающей [x1..x2].
						//     ВАЖНО: если встретили ЧЕРТУ — оставляем границу НА этой строке.
						int denBottom = y + 1;

						while (denBottom < rows) {
							if (row_has_overlapping_bar(denBottom, x1, x2)) {
								break;
							}

							const bool allSpace = Model::Geometry::IsRowRangeAllSpace(
								grid,
								denBottom,
								x1,
								x2
							);

							if (allSpace) {
								--denBottom;
								break;
							}

							++denBottom;
						}

						if (denBottom >= rows) {
							denBottom = rows - 1;
						}

						const int denTop = y + 1;

						// --- собрать регионы строго в X = [x1..x2]
						Model::Region numRegion{
							Model::SpanY{ numTop, numBottom },
							Model::SpanX{ x1, x2 }
						};

						Model::Region denRegion{
							Model::SpanY{ denTop, denBottom },
							Model::SpanX{ x1, x2 }
						};

						bars.push_back(FractionBar{
							.y = y,
							.x1 = x1,
							.x2 = x2,
							.numRegion = numRegion,
							.denRegion = denRegion,
							});
					}

					return bars;
				}

			private:
				bool IsHorizontalBarChar(char c) const {
					return c == '-' || c == '=' || c == '─' || c == '━';
				}
			};
		}
	}
}
