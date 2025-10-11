#pragma once
#include "../Detect/IFeatureDetector.h"
#include "../Layout/LayoutGraph.h"
#include "../Model/Geometry.h"
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
				int    minTextRunLen = 1;         // минимальная длина текстового сегмента
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
					this->AddTextRuns_(g, grid, features);   // шаг 1
					this->AddFeatureNodes_(g, features);     // шаг 2
					this->WireSpatialRelations_(g, grid);    // шаг 3
					return g;
				}

			private:
				// ↓↓↓ служебные методы ↓↓↓

				// Проверяем, вся ли область состоит из '='.
				bool IsAllDashes_(const Model::AsciiGrid& grid, const Model::Region& r) const {
					for (int y = r.rows.y1; y <= r.rows.y2; ++y)
						for (int x = r.cols.x1; x <= r.cols.x2; ++x)
							if (grid.At(y, x) != '=') return false;
					return true;
				}

				// Проверяем, есть ли хотя бы одна пустая строка между областями.
				// Если области вплотную — тоже считаем, что разрыв допустим.
				bool HasEmptyRowBetween_(
					const Model::AsciiGrid& grid,
					const Model::Region& upper,
					const Model::Region& lower
				) const {
					if (lower.rows.y1 - upper.rows.y2 <= 1)
						return true;
					for (int y = upper.rows.y2 + 1; y <= lower.rows.y1 - 1; ++y)
						if (Model::Geometry::IsRowRangeAllSpace(grid, y, lower.cols.x1, lower.cols.x2))
							return true;
					return false;
				}

				// Сохраняем координаты всех bar-фич, чтобы не создавать
				// дублирующие text-run'ы из тех же «====».
				std::unordered_set<std::uint64_t> BuildBarKeySet_(
					const std::vector<Detect::FeatureRegion>& features
				) const {
					std::unordered_set<std::uint64_t> keys{};
					for (const auto& f : features)
						if (f.featureKind == "bar") {
							const auto& r = f.region;
							const std::uint64_t key =
								(static_cast<std::uint64_t>(r.rows.y1) << 48) ^
								(static_cast<std::uint64_t>(r.rows.y2) << 32) ^
								(static_cast<std::uint64_t>(r.cols.x1) << 16) ^
								(static_cast<std::uint64_t>(r.cols.x2));
							keys.insert(key);
						}
					return keys;
				}

				bool IsExactBarRegion_(
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

				// 1️⃣ Создание текстовых узлов (text-run)
				void AddTextRuns_(
					LayoutGraph& outGraph,
					const Model::AsciiGrid& grid,
					const std::vector<Detect::FeatureRegion>& features
				) const {
					const auto barKeys = this->BuildBarKeySet_(features);
					for (int y = 0; y < grid.Height(); ++y) {
						int x = 0;
						while (x < grid.Width()) {
							while (x < grid.Width() && grid.At(y, x) == ' ') ++x;
							if (x >= grid.Width()) break;
							const int xStart = x;
							while (x < grid.Width() && grid.At(y, x) != ' ') ++x;
							const int xEnd = x - 1;
							const int len = xEnd - xStart + 1;
							if (len < this->options.minTextRunLen) continue;

							Model::Region r{ Model::SpanY{ y, y }, Model::SpanX{ xStart, xEnd } };
							// Пропускаем сегменты, которые совпадают с bar
							if (this->IsAllDashes_(grid, r) && this->IsExactBarRegion_(barKeys, r))
								continue;

							outGraph.AddNode(r, "text-run");
						}
					}
				}

				// 2️⃣ Добавление узлов-фич (bar, sqrt, bracket …)
				void AddFeatureNodes_(
					LayoutGraph& outGraph,
					const std::vector<Detect::FeatureRegion>& features
				) const {
					for (const auto& f : features)
						outGraph.AddNode(f.region, f.featureKind);
				}

				// 3️⃣ Построение связей bar ↔ text-run
				void WireSpatialRelations_(
					LayoutGraph& outGraph,
					const Model::AsciiGrid& grid
				) const {
					const auto& nodesRef = outGraph.Nodes();
					const std::size_t n = nodesRef.size();

					auto add_edge_once = [&](Layout::NodeId u, Layout::NodeId v, RelKind k) {
						for (const auto& e : outGraph.Edges())
							if (e.a == u && e.b == v && e.kind == k) return;
						outGraph.AddEdge(u, v, k);
						};

					for (std::size_t i = 0; i < n; ++i) {
						const auto& bar = nodesRef[i];
						if (bar.role != "bar") continue;

						for (std::size_t j = 0; j < n; ++j) {
							if (i == j) continue;
							const auto& txt = nodesRef[j];
							if (txt.role != "text-run") continue;

							// Проверяем, пересекаются ли по X достаточно сильно
							const int overlapX = bar.region.OverlapX(txt.region);
							const int baseBar = std::max(1, bar.region.Width());
							const int baseTxt = std::max(1, txt.region.Width());
							const double relBar = double(overlapX) / baseBar;
							const double relTxt = double(overlapX) / baseTxt;
							if (relBar < this->options.minRelativeXOverlap &&
								relTxt < this->options.minRelativeXOverlap)
								continue;

							// Добавляем ориентированную связь: (верх) --Above--> (низ)
							if (txt.region.rows.y2 < bar.region.rows.y1) {
								if (this->HasEmptyRowBetween_(grid, txt.region, bar.region))
									add_edge_once(txt.id, bar.id, RelKind::Above);
							}
							else if (bar.region.rows.y2 < txt.region.rows.y1) {
								if (this->HasEmptyRowBetween_(grid, bar.region, txt.region))
									add_edge_once(bar.id, txt.id, RelKind::Above);
							}
						}
					}
				}

			private:
				LayoutBuilderOptions options;
			};
		}
	}
}
