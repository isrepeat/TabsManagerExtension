#pragma once
#include "Grid.h"
#include <algorithm>
#include <optional>

namespace AsciiMathParser {
	namespace Core {
		namespace Model {
			//
			// ░ Geometry
			// ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
			//
			// Вспомогательный набор статических утилит для работы с AsciiGrid.
			class Geometry {
			public:
				// Проверяет, что вся строка `y` в диапазоне [x1, x2] состоит только из пробелов.
				// Используется для нахождения "пустых" разделительных строк между элементами дроби
				// и при вычислении TightBBoxAbove / TightBBoxBelow.
				static bool IsRowRangeAllSpace(
					const Model::AsciiGrid& grid,
					int y,
					int x1,
					int x2
				) {
					for (int x = x1; x <= x2; ++x) {
						const char ch = grid.At(y, x);
						if (ch != ' ') {
							return false;
						}
					}
					return true;
				}

				// Возвращает минимальный ограничивающий прямоугольник (bbox)
				// непустых символов над указанным горизонтальным "bandRegion".
				// Алгоритм:
				//   1) двигаемся вверх от bandRegion.rows.y1, пока не встретим пустую строку.
				//   2) ограничиваем область [yTop, yBottom] между первой непустой и последней.
				//   3) вычисляем minX / maxX по непустым символам.
				// Если непустых строк нет — возвращает std::nullopt.
				static std::optional<Model::Region> TightBBoxAbove(
					const Model::AsciiGrid& grid,
					const Model::Region& bandRegion
				) {
					const int yBottom = bandRegion.rows.y1 - 1;
					if (yBottom < 0) {
						return std::nullopt;
					}

					int y = yBottom;
					while (y >= 0) {
						if (Geometry::IsRowRangeAllSpace(grid, y, bandRegion.cols.x1, bandRegion.cols.x2)) {
							break;
						}
						--y;
					}
					const int yTop = y + 1;
					if (yTop > yBottom) {
						return std::nullopt;
					}

					int minX = std::numeric_limits<int>::max();
					int maxX = std::numeric_limits<int>::min();

					for (int yy = yTop; yy <= yBottom; ++yy) {
						for (int xx = bandRegion.cols.x1; xx <= bandRegion.cols.x2; ++xx) {
							const char ch = grid.At(yy, xx);
							if (ch != ' ') {
								if (xx < minX) {
									minX = xx;
								}
								if (xx > maxX) {
									maxX = xx;
								}
							}
						}
					}

					if (maxX < minX) {
						return std::nullopt;
					}

					Model::Region res{
						Model::SpanY{ yTop, yBottom },
						Model::SpanX{ minX, maxX }
					};
					return res;
				}

				// Аналог TightBBoxAbove, но ищет область непустых символов под горизонтальной полосой.
				// Идём вниз, пока не встретим пустую строку, и строим плотный bbox для символов в диапазоне X.
				static std::optional<Model::Region> TightBBoxBelow(
					const Model::AsciiGrid& grid,
					const Model::Region& bandRegion
				) {
					const int yTop = bandRegion.rows.y2 + 1;
					if (yTop >= grid.Height()) {
						return std::nullopt;
					}

					int y = yTop;
					while (y < grid.Height()) {
						if (Geometry::IsRowRangeAllSpace(grid, y, bandRegion.cols.x1, bandRegion.cols.x2)) {
							break;
						}
						++y;
					}
					const int yBottom = y - 1;
					if (yBottom < yTop) {
						return std::nullopt;
					}

					int minX = std::numeric_limits<int>::max();
					int maxX = std::numeric_limits<int>::min();

					for (int yy = yTop; yy <= yBottom; ++yy) {
						for (int xx = bandRegion.cols.x1; xx <= bandRegion.cols.x2; ++xx) {
							const char ch = grid.At(yy, xx);
							if (ch != ' ') {
								if (xx < minX) {
									minX = xx;
								}
								if (xx > maxX) {
									maxX = xx;
								}
							}
						}
					}

					if (maxX < minX) {
						return std::nullopt;
					}

					Model::Region res{
						Model::SpanY{ yTop, yBottom },
						Model::SpanX{ minX, maxX }
					};
					return res;
				}

				// Проверяет, пересекаются ли регионы по оси X.
				// Используется при определении связей Above / Below между элементами.
				static bool OverlapsX(
					const Model::Region& a,
					const Model::Region& b
				) {
					return (a.OverlapX(b) > 0);
				}

				// Проверяет, что `a` расположен строго выше `b` (не касается, не перекрывает).
				// Условие: нижняя граница `a` меньше верхней границы `b`.
				static bool IsAbove(
					const Model::Region& a,
					const Model::Region& b
				) {
					if (a.rows.y2 < b.rows.y1) {
						return true;
					}
					return false;
				}

				// Проверяет, что `a` расположен строго ниже `b`.
				// Условие: верхняя граница `a` больше нижней границы `b`.
				static bool IsBelow(
					const Model::Region& a,
					const Model::Region& b
				) {
					if (a.rows.y1 > b.rows.y2) {
						return true;
					}
					return false;
				}
			};
		}
	}
}