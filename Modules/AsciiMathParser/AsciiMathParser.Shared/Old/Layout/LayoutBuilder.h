#pragma once
#include "../Detect/IFeatureDetector.h"
#include "../Model/Geometry.h"
#include "RelationBuilderContext.h"
#include "BarRelationBuilder.h"
#include "LayoutGraph.h"

#include <unordered_set>
#include <vector>
#include <memory>

//
// ░ Description
// ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
//
// В традиционной записи выражений (на бумаге или в LaTeX) операторы
// располагаются не только по горизонтали, но и по вертикали или даже
// диагонально. Поэтому парсинг таких формул через обычную линейную
// грамматику невозможен.
//
// Здесь применяется обобщённая геометрическая модель:
//
//   ▫ LayoutGraph описывает *пространственные отношения* между фрагментами:
//       Above / Below / Inside / LeftOf / RightOf / Adjacent.
//   ▫ LayoutBuilder строит этот граф из ASCII-сетки.
//   ▫ Далее специализированные "Builder"-ы (FractionBuilder, LinearBuilder, …)
//     анализируют связи и создают соответствующие узлы AST.
//
// ░ Типы операторов и их пространственная ориентация:
//
//   1. Линейные операторы  ( + , - , * , / )
//        → горизонтальные связи: LeftOf / RightOf
//        → пример:  a + b * c
//
//   2. Вертикальные операторы  ( черта дроби )
//        → вертикальные связи: Above / Below
//        → пример:
//                  a + b
//             -------------
//                    z
//
//   3. Вложенные операторы  ( √ , ∫ , Σ , скобки )
//        → вложенные связи: Inside
//        → пример:  √(x + y)
//
//   4. Диагональные операторы  ( степени и индексы )
//        → наклонные связи: Above-Right / Below-Right (вариации Above/Below)
//        → пример:  x²   или   xₙ
//
// ░ Зачем всё это:
//
//   • LayoutGraph отделяет *геометрию выражения* от *семантики операторов*.
//     Он не знает, что такое «дробь» или «умножение» — только «что над чем».
//   • Это делает систему расширяемой: чтобы добавить новый тип оператора,
//     нужно лишь реализовать новый детектор и билдер, не трогая остальной код.
//   • Один и тот же LayoutGraph может быть интерпретирован разными билдерами,
//     каждый из которых смотрит только на свои типы отношений.
//
// Таким образом, LayoutGraph — это универсальный "каркас" пространства формулы,
// а LayoutBuilder — его геометрический конструктор.
//

namespace AsciiMathParser {
	namespace Core {
		namespace Layout {
			struct LayoutBuilderOptions {
				double minRelativeXOverlap = 0.5; // требуемая доля перекрытия по X
				int minTextRunLen = 1;            // минимальная длина текстового сегмента
			};

			//
			// ░ LayoutBuilder
			// ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
			//
			// Строит LayoutGraph в 3 этапа:
			//   1) Разбивает ASCII-текст на text-run'ы.
			//   2) Добавляет фичи, найденные детекторами (bar и др.).
			//   3) Строит рёбра отношений между узлами (Above/Below и т.п.).
			//
			class LayoutBuilder {
			public:
				explicit LayoutBuilder(const LayoutBuilderOptions& options)
					: options{ options } {
				}

				LayoutGraph Build(
					const Model::AsciiGrid& grid,
					const std::vector<Detect::FeatureRegion>& features
				) const {
					LayoutGraph g{};
					this->AddTextRuns(g, grid, features);   // шаг 1
					this->AddFeatureNodes(g, features);     // шаг 2
					this->WireSpatialRelations(g, grid);    // шаг 3
					return g;
				}

			private:
				// 1️. Создание текстовых узлов (text-run)
				void AddTextRuns(
					LayoutGraph& outGraph,
					const Model::AsciiGrid& grid,
					const std::vector<Detect::FeatureRegion>& features
				) const {
					const auto barKeys = this->BuildBarKeySet(features);
					for (int y = 0; y < grid.Height(); ++y) {
						int x = 0;
						while (x < grid.Width()) {
							while (x < grid.Width() && grid.At(y, x) == ' ') {
								++x;
							}
							if (x >= grid.Width()) {
								break;
							}

							const int xStart = x;
							while (x < grid.Width() && grid.At(y, x) != ' ') {
								++x;
							}

							const int xEnd = x - 1;
							const int len = xEnd - xStart + 1;
							if (len < this->options.minTextRunLen) {
								continue;
							}

							Model::Region r{ 
								Model::SpanY{ y, y },
								Model::SpanX{ xStart, xEnd }
							};

							// Пропускаем сегменты, которые совпадают с bar
							if (this->IsAllDashes(grid, r) && this->IsExactBarRegion(barKeys, r)) {
								continue;
							}
							outGraph.AddNode(grid, r, "text-run");
						}
					}
				}


				// 2️. Добавление узлов-фич (bar, sqrt, bracket …)
				void AddFeatureNodes(
					LayoutGraph& outGraph,
					const std::vector<Detect::FeatureRegion>& features
				) const {
					for (const auto& f : features) {
						outGraph.AddNode(f.region, f.featureKind);
					}
				}

				// 3️. Построение связей bar ↔ text-run
				void WireSpatialRelations(
					LayoutGraph& outGraph,
					const Model::AsciiGrid& grid
				) const {
					RelationBuilderContext ctx{ outGraph, grid };

					BarRelationBuilder barBuilder{ this->options.minRelativeXOverlap };
					barBuilder.Wire(ctx);

					// В будущем можно добавить:
					// SqrtRelationBuilder sqrtBuilder{};
					// sqrtBuilder.Wire(ctx);


					//const auto& nodesRef = outGraph.Nodes();
					//const std::size_t n = nodesRef.size();

					//auto add_edge_once = [&](Layout::NodeId u, Layout::NodeId v, RelKind k) {
					//	for (const auto& e : outGraph.Edges()) {
					//		if (e.a == u && e.b == v && e.kind == k) {
					//			return;
					//		}
					//	}
					//	outGraph.AddEdge(u, v, k);
					//	};

					//// Соберём список баров для проверки «чистого коридора»
					//const auto barRegions = this->CollectBarRegions(outGraph);

					//for (std::size_t i = 0; i < n; ++i) {
					//	const auto& bar = nodesRef[i];
					//	if (bar.role != "bar") {
					//		continue;
					//	}

					//	for (std::size_t j = 0; j < n; ++j) {
					//		if (i == j) {
					//			continue;
					//		}

					//		const auto& other = nodesRef[j];
					//		if (other.role != "text-run" && other.role != "bar") {
					//			continue;
					//		}

					//		// Перекрытие по X должно быть достаточным
					//		const int overlapX = bar.region.OverlapX(other.region);
					//		const int baseBar = std::max(1, bar.region.Width());
					//		const int baseOth = std::max(1, other.region.Width());
					//		const double relBar = double(overlapX) / baseBar;
					//		const double relOth = double(overlapX) / baseOth;

					//		if (relBar < this->options.minRelativeXOverlap &&
					//			relOth < this->options.minRelativeXOverlap) {
					//			continue;
					//		}

					//		// Ориентация: (верх) --Above--> (низ)
					//		if (other.region.rows.y2 < bar.region.rows.y1) {
					//			// other над bar
					//			const bool ok =
					//				(other.role == "bar")
					//				? true // bar→bar: разрешаем без коридора (вложенные дроби)
					//				: this->HasClearVerticalCorridor(grid, other.region, bar.region, barRegions);

					//			if (ok) {
					//				add_edge_once(other.id, bar.id, RelKind::Above);
					//			}
					//		}
					//		else if (bar.region.rows.y2 < other.region.rows.y1) {
					//			// bar над other
					//			const bool ok =
					//				(other.role == "bar")
					//				? true // bar→bar: разрешаем без коридора (вложенные дроби)
					//				: this->HasClearVerticalCorridor(grid, bar.region, other.region, barRegions);

					//			if (ok) {
					//				add_edge_once(bar.id, other.id, RelKind::Above);
					//			}
					//		}
					//	}
					//}
				}

				//// Быстрый список всех bar-регионов (по узлам графа)
				//std::vector<Model::Region> CollectBarRegions(
				//	const LayoutGraph& g
				//) const {
				//	std::vector<Model::Region> bars{};
				//	for (const auto& n : g.Nodes()) {
				//		if (n.role == "bar") {
				//			bars.push_back(n.region);
				//		}
				//	}
				//	return bars;
				//}

				//// Принадлежит ли клетка (y,x) какому-либо bar-региону
				//bool IsCellInAnyBar(
				//	const std::vector<Model::Region>& bars,
				//	const int y,
				//	const int x
				//) const {
				//	for (const auto& r : bars) {
				//		if (r.rows.y1 <= y && y <= r.rows.y2) {
				//			if (r.cols.x1 <= x && x <= r.cols.x2) {
				//				return true;
				//			}
				//		}
				//	}
				//	return false;
				//}

				//// «Чистый вертикальный коридор» для пары (верх → низ) по X-перекрытию.
				//// Разрешаем ' ' и '=' (если '=' принадлежит какому-то bar).
				//bool HasClearVerticalCorridor(
				//	const Model::AsciiGrid& grid,
				//	const Model::Region& upper,
				//	const Model::Region& lower,
				//	const std::vector<Model::Region>& bars
				//) const {
				//	const int overlapX1 = std::max(upper.cols.x1, lower.cols.x1);
				//	const int overlapX2 = std::min(upper.cols.x2, lower.cols.x2);
				//	if (overlapX2 < overlapX1) {
				//		return false; // нет реального перекрытия по X
				//	}

				//	// Вплотную по Y — считаем коридор «чистым»
				//	if (lower.rows.y1 - upper.rows.y2 <= 1) {
				//		return true;
				//	}

				//	// Проверяем вертикальный коридор между верхним и нижним регионами.
				//	// Сканируем только по тем колонкам (x), где области реально пересекаются.
				//	for (int x = overlapX1; x <= overlapX2; ++x) {
				//		for (int y = upper.rows.y2 + 1; y <= lower.rows.y1 - 1; ++y) {
				//			const char c = grid.At(y, x);
				//			if (c == ' ') {
				//				continue;
				//			}
				//			if (c == '=') {
				//				if (this->IsCellInAnyBar(bars, y, x)) {
				//					continue; // '=' внутри любого bar — это допустимо
				//				}
				//			}
				//			return false; // встретили «посторонний» символ → коридор «грязный»
				//		}
				//	}
				//	return true;
				//}


				// Проверяем, вся ли область состоит из '='.
				bool IsAllDashes(const Model::AsciiGrid& grid, const Model::Region& r) const {
					for (int y = r.rows.y1; y <= r.rows.y2; ++y) {
						for (int x = r.cols.x1; x <= r.cols.x2; ++x) {
							if (grid.At(y, x) != '=') {
								return false;
							}
						}
					}
					return true;
				}


				// Сохраняем координаты всех bar-фич, чтобы не создавать
				// дублирующие text-run'ы из тех же «====».
				std::unordered_set<std::uint64_t> BuildBarKeySet(
					const std::vector<Detect::FeatureRegion>& features
				) const {
					std::unordered_set<std::uint64_t> keys{};
					for (const auto& f : features) {
						if (f.featureKind == "bar") {
							const auto& r = f.region;
							const std::uint64_t key =
								(static_cast<std::uint64_t>(r.rows.y1) << 48) ^
								(static_cast<std::uint64_t>(r.rows.y2) << 32) ^
								(static_cast<std::uint64_t>(r.cols.x1) << 16) ^
								(static_cast<std::uint64_t>(r.cols.x2));
							keys.insert(key);
						}
					}
					return keys;
				}

				bool IsExactBarRegion(
					const std::unordered_set<std::uint64_t>& barKeys,
					const Model::Region& r
				) const {
					const std::uint64_t key =
						(static_cast<std::uint64_t>(r.rows.y1) << 48) ^
						(static_cast<std::uint64_t>(r.rows.y2) << 32) ^
						(static_cast<std::uint64_t>(r.cols.x1) << 16) ^
						(static_cast<std::uint64_t>(r.cols.x2));
					return (barKeys.find(key) != barKeys.end());
				}

			private:
				LayoutBuilderOptions options;
			};
		}
	}
}
