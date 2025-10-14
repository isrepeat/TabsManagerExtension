#pragma once
#include "../Model/Geometry.h"
#include "../Model/Grid.h"
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		namespace Detect {
			struct FractionBar {
				Model::RowRegion bar;
				Model::Region numRegion;
				Model::Region denRegion;
			};

			class FractionBarDetector {
			public:
				std::vector<FractionBar> DetectBars(const Model::AsciiGrid& grid) const {
					std::vector<FractionBar> bars;

					const int rows = grid.Height();
					const int cols = grid.Width();

					// 1) Собираем все горизонтальные «полосы» как RowRegion
					std::vector<Model::RowRegion> rawBars{};
					rawBars.reserve(static_cast<std::size_t>(rows));

					for (int y = 0; y < rows; ++y) {
						int x = 0;

						while (x < cols) {
							// Пропускаем не-черточные символы
							while (x < cols && !this->IsHorizontalBarChar(grid.At(x, y))) {
								++x;
							}
							if (x >= cols) {
								break;
							}

							const int startX = x;

							// Идём вправо по непрерывной черте
							while (x < cols && this->IsHorizontalBarChar(grid.At(x, y))) {
								++x;
							}
							const int endX = x - 1;

							// Проверяем наличие контента сверху и снизу
							bool hasAbove = false;
							bool hasBelow = false;

							if (y > 0) {
								if (!Model::Geometry::IsRowRangeAllSpace(grid, y - 1, startX, endX)) {
									hasAbove = true;
								}
							}
							if (y + 1 < rows) {
								if (!Model::Geometry::IsRowRangeAllSpace(grid, y + 1, startX, endX)) {
									hasBelow = true;
								}
							}

							if (hasAbove && hasBelow) {
								rawBars.push_back(Model::RowRegion{
									y,
									Model::SpanX{ startX, endX }
									});
							}
						}
					}

					if (rawBars.empty()) {
						return bars;
					}

					//// 2) Проверка пересечений между строками
					//auto rowHasOverlap = [&](const int y, const int x1, const int x2) -> bool {
					//	if (y < 0 || y >= rows) {
					//		return false;
					//	}
					//	for (const auto& rb : rawBars) {
					//		if (rb.y != y) {
					//			continue;
					//		}
					//		const int L = std::max(rb.cols.x1, x1);
					//		const int R = std::min(rb.cols.x2, x2);
					//		if (L <= R) {
					//			return true;
					//		}
					//	}
					//	return false;
					//	};

					//// 3) Для каждой черты определяем регионы числителя и знаменателя
					//for (const auto& rb : rawBars) {
					//	const int y = rb.y;
					//	const int x1 = rb.cols.x1;
					//	const int x2 = rb.cols.x2;

					//	int numTop = y - 1;
					//	while (numTop >= 0) {
					//		if (rowHasOverlap(numTop, x1, x2)) {
					//			break;
					//		}
					//		if (Model::Geometry::IsRowRangeAllSpace(grid, numTop, x1, x2)) {
					//			++numTop;
					//			break;
					//		}
					//		--numTop;
					//	}
					//	if (numTop < 0) {
					//		numTop = 0;
					//	}
					//	const int numBottom = y - 1;

					//	int denBottom = y + 1;
					//	while (denBottom < rows) {
					//		if (rowHasOverlap(denBottom, x1, x2)) {
					//			break;
					//		}
					//		if (Model::Geometry::IsRowRangeAllSpace(grid, denBottom, x1, x2)) {
					//			--denBottom;
					//			break;
					//		}
					//		++denBottom;
					//	}
					//	if (denBottom >= rows) {
					//		denBottom = rows - 1;
					//	}
					//	const int denTop = y + 1;

					//	Model::RowRegion barRegion{ y, Model::SpanX{ x1, x2 } };
					//	Model::Region numRegion{
					//		Model::SpanX{ x1, x2 },
					//		Model::SpanY{ numTop, numBottom }
					//	};
					//	Model::Region denRegion{
					//		Model::SpanX{ x1, x2 },
					//		Model::SpanY{ denTop, denBottom }
					//	};

					//	bars.push_back(FractionBar{
					//		.bar = barRegion,
					//		.numRegion = numRegion,
					//		.denRegion = denRegion
					//		});
					//}


					// 3) Для каждой черты строим регионы
					for (const auto& barRowRegion : rawBars) {
						const int numTop = FindNumTopInclusive(grid, rawBars, rows, barRowRegion);
						const int numBottom = barRowRegion.y - 1;

						const int denTop = barRowRegion.y + 1;
						const int denBottom = FindDenBottomInclusive(grid, rawBars, rows, barRowRegion);

						Model::Region numRegion = Model::Region{
							barRowRegion.cols,
							Model::SpanY{ numTop, numBottom }
						};

						Model::Region denRegion = Model::Region{
							barRowRegion.cols,
							Model::SpanY{ denTop, denBottom }
						};

						bars.push_back(FractionBar{
							.bar = barRowRegion,
							.numRegion = numRegion,
							.denRegion = denRegion
							});
					}

					return bars;
				}

			private:
				// Есть ли на строке y другая черта, пересекающаяся по X с текущим окном бара?
				static bool RowHasOverlappingBar(
					const std::vector<Model::RowRegion>& allBars,
					const int                            y,
					const Model::RowRegion& currentBar
				) {
					const Model::SpanX& window = currentBar.cols;

					for (const auto& rb : allBars) {
						if (rb.y != y) {
							continue;
						}
						// пропускаем сам бар (если сравниваем сам с собой)
						if (&rb == &currentBar) {
							continue;
						}
						const int L = std::max(rb.cols.x1, window.x1);
						const int R = std::min(rb.cols.x2, window.x2);
						if (L <= R) {
							return true;
						}
					}
					return false;
				}

				static int FindNumTopInclusive(
					const Model::AsciiGrid& grid,
					const std::vector<Model::RowRegion>& allBars,
					const int                            totalRows,
					const Model::RowRegion& bar
				) {
					int y = bar.y - 1;

					while (y >= 0) {
						if (RowHasOverlappingBar(allBars, y, bar)) {
							break;
						}
						if (Model::Geometry::IsRowRangeAllSpace(grid, y, bar.cols.x1, bar.cols.x2)) {
							++y;
							break;
						}
						--y;
					}

					return std::max(0, y);
				}


				static int FindDenBottomInclusive(
					const Model::AsciiGrid& grid,
					const std::vector<Model::RowRegion>& allBars,
					const int                            totalRows,
					const Model::RowRegion& bar
				) {
					int y = bar.y + 1;

					while (y < totalRows) {
						if (RowHasOverlappingBar(allBars, y, bar)) {
							break;
						}
						if (Model::Geometry::IsRowRangeAllSpace(grid, y, bar.cols.x1, bar.cols.x2)) {
							--y;
							break;
						}
						++y;
					}

					return std::min(totalRows - 1, y);
				}


				bool IsHorizontalBarChar(char c) const {
					return c == '-' || c == '=' || c == '─' || c == '━';
				}
			};
		}
	}
}
