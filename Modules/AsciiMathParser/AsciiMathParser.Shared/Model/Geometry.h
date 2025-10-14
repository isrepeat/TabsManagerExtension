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
				static Model::RowRegion MakeRowRegion(
					const int y,
					const int x1,
					const int x2
				) {
					return Model::RowRegion{
						y,
						Model::SpanX{ x1, x2 }
					};
				}

				static Model::ColRegion MakeColRegion(
					const int x,
					const int y1,
					const int y2
				) {
					return Model::ColRegion{
						x,
						Model::SpanY{ y1, y2 }
					};
				}


				// Объединение двух регионов (bounding box)
				static Model::Region UnionRegion(
					const Model::Region& a,
					const Model::Region& b
				) {
					Model::Region r{
						Model::SpanX{
							std::min(a.cols.x1, b.cols.x1),
							std::max(a.cols.x2, b.cols.x2)
					},
						Model::SpanY{
							std::min(a.rows.y1, b.rows.y1),
							std::max(a.rows.y2, b.rows.y2)
					}
					};
					return r;
				}


				template<typename... TRest>
				static Model::Region UnionRegions(
					const Model::Region& r0,
					const TRest&... rest
				) {
					static_assert((std::is_same_v<std::decay_t<TRest>, Model::Region> && ...),
						"All arguments must be Model::Region");

					std::optional<Model::Region> acc{ r0 };

					if constexpr (sizeof...(rest) > 0) {
						(acc.emplace(Geometry::UnionRegion(*acc, rest)), ...);
					}
					return acc.value();
				}


				// Унион регионов произвольного набора "узлов" с методом GetRegion()
				// Примерно: UnionRegionsOfNodes(group.nodes)
				template<class TContainer>
				static Model::Region UnionRegionsOfNodes(
					const TContainer& nodes
				) {
					std::optional<Model::Region> acc{};
					for (const auto& it : nodes) {
						// поддержка как T, так и unique_ptr<T>/shared_ptr<T>
						const auto& nodeRef = (it ? *it : *it); // для сырых T компилятор сам упростит
						Model::Region r = nodeRef.GetRegion();
						if (acc.has_value()) {
							acc.emplace(Geometry::UnionRegion(*acc, r));
						}
						else {
							acc.emplace(r);
						}
					}
					return acc.value_or(Model::Region{});
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
						const char ch = grid.At(x, y);
						if (ch != ' ') {
							return false;
						}
					}
					return true;
				}
			};
		}
	}
}