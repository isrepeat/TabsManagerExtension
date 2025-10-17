#pragma once
#include "../Model/Geometry.h"
#include "../Model/Grid.h"
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		namespace Parse {
			struct FractionBar {
				Model::RowRegion barRegion;
				Model::Region numRegion;
				Model::Region denRegion;
			};

			class FractionBarDetector {
			public:
				FractionBarDetector(const Model::AsciiGrid& grid)
					: grid{ grid } {
				}


				std::vector<FractionBar> DetectBars() const {
					std::vector<FractionBar> bars;

					const int rows = this->grid.Height();
					const int cols = this->grid.Width();

					// 1) Собираем все горизонтальные «полосы» как RowRegion
					std::vector<Model::RowRegion> rawBars{};
					rawBars.reserve(static_cast<std::size_t>(rows));

					for (int y = 0; y < rows; ++y) {
						int x = 0;

						while (x < cols) {
							// Пропускаем не-черточные символы
							while (x < cols && !this->IsHorizontalBarChar(this->grid.At(x, y))) {
								++x;
							}
							if (x >= cols) {
								break;
							}

							const int startX = x;

							// Идём вправо по непрерывной черте
							while (x < cols && this->IsHorizontalBarChar(this->grid.At(x, y))) {
								++x;
							}
							const int endX = x - 1;

							// Проверяем наличие контента сверху и снизу
							bool hasAbove = false;
							bool hasBelow = false;

							if (y > 0) {
								if (!Model::Geometry::IsRowRangeAllSpace(this->grid, y - 1, startX, endX)) {
									hasAbove = true;
								}
							}
							if (y + 1 < rows) {
								if (!Model::Geometry::IsRowRangeAllSpace(this->grid, y + 1, startX, endX)) {
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

					// 3) Для каждой черты строим регионы
					for (const auto& barRowRegion : rawBars) {
						const int numTop = this->FindNumTopInclusive(rows, barRowRegion, rawBars);
						const int numBottom = barRowRegion.y - 1;

						const int denTop = barRowRegion.y + 1;
						const int denBottom = this->FindDenBottomInclusive(rows, barRowRegion, rawBars);

						Model::Region numRegion = Model::Region{
							barRowRegion.cols,
							Model::SpanY{ numTop, numBottom }
						};

						Model::Region denRegion = Model::Region{
							barRowRegion.cols,
							Model::SpanY{ denTop, denBottom }
						};

						bars.push_back(FractionBar{
							.barRegion = barRowRegion,
							.numRegion = numRegion,
							.denRegion = denRegion
							});
					}

					return bars;
				}

			private:
				// Есть ли на строке y другая черта, пересекающаяся по X с текущим окном бара?
				bool RowHasOverlappingBar(
					const int y,
					const Model::RowRegion& bar,
					const std::vector<Model::RowRegion>& allBars
				) const {
					const Model::SpanX& window = bar.cols;

					for (const auto& rb : allBars) {
						if (rb.y != y) {
							continue;
						}
						// пропускаем сам бар (если сравниваем сам с собой)
						if (&rb == &bar) {
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


				int FindNumTopInclusive(
					const int totalRows,
					const Model::RowRegion& bar,
					const std::vector<Model::RowRegion>& allBars
				) const {
					int y = bar.y - 1;

					while (y >= 0) {
						if (this->RowHasOverlappingBar(y, bar, allBars)) {
							break;
						}
						if (Model::Geometry::IsRowRangeAllSpace(this->grid, y, bar.cols.x1, bar.cols.x2)) {
							++y;
							break;
						}
						--y;
					}

					return std::max(0, y);
				}


				int FindDenBottomInclusive(
					const int totalRows,
					const Model::RowRegion& bar,
					const std::vector<Model::RowRegion>& allBars
				) const {
					int y = bar.y + 1;

					while (y < totalRows) {
						if (this->RowHasOverlappingBar(y, bar, allBars)) {
							break;
						}
						if (Model::Geometry::IsRowRangeAllSpace(this->grid, y, bar.cols.x1, bar.cols.x2)) {
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

				private:
					const Model::AsciiGrid& grid;
			};
		}
	}
}
