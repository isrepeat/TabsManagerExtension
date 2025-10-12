#pragma once
#include "../Detect/IFeatureDetector.h"
#include "RelationBuilderContext.h"
#include "LayoutGraph.h"

#include <unordered_set>
#include <vector>
#include <memory>

namespace AsciiMathParser {
	namespace Core {
		namespace Layout {
			class BarRelationBuilder {
			public:
				explicit BarRelationBuilder(double minRelativeXOverlap)
					: minRelativeXOverlap{ minRelativeXOverlap } {
				}

				void Wire(RelationBuilderContext& ctx) const {
					const auto barRegions = this->CollectBarRegions(ctx.graph);

					for (const LayoutNode* bar : ctx.byRole["bar"]) {
						for (const auto& node : ctx.graph.Nodes()) {
							if (node.id == bar->id) {
								continue;
							}
							if (node.role != "text-run" && node.role != "bar"
								&& node.role != "sqrt" && node.role != "bracket") {
								continue;
							}

							const int overlapX = bar->region.OverlapX(node.region);
							const int baseBar = std::max(1, bar->region.Width());
							const int baseNode = std::max(1, node.region.Width());
							const double relBar = static_cast<double>(overlapX) / baseBar;
							const double relNode = static_cast<double>(overlapX) / baseNode;

							if (relBar < this->minRelativeXOverlap &&
								relNode < this->minRelativeXOverlap) {
								continue;
							}

							auto add_edge_once = [&](Layout::NodeId u, Layout::NodeId v, RelKind k) {
								for (const auto& e : ctx.graph.Edges()) {
									if (e.a == u && e.b == v && e.kind == k) {
										return;
									}
								}
								ctx.graph.AddEdge(u, v, k);
								};

							// (верх) --Above--> (низ)
							if (node.region.rows.y2 < bar->region.rows.y1) {
								const bool ok =
									(node.role == "bar")
									? true
									: this->HasClearVerticalCorridor(
										ctx.grid, node.region, bar->region, barRegions);

								if (ok) {
									add_edge_once(node.id, bar->id, RelKind::Above);
								}
							}
							else if (bar->region.rows.y2 < node.region.rows.y1) {
								const bool ok =
									(node.role == "bar")
									? true
									: this->HasClearVerticalCorridor(
										ctx.grid, bar->region, node.region, barRegions);

								if (ok) {
									add_edge_once(bar->id, node.id, RelKind::Above);
								}
							}
						}
					}
				}

			private:
				// «Чистый вертикальный коридор» для пары (верх → низ) по X-перекрытию.
				// Разрешаем ' ' и '=' (если '=' принадлежит какому-то bar).
				bool HasClearVerticalCorridor(
					const Model::AsciiGrid& grid,
					const Model::Region& upper,
					const Model::Region& lower,
					const std::vector<Model::Region>& bars
				) const {
					const int overlapX1 = std::max(upper.cols.x1, lower.cols.x1);
					const int overlapX2 = std::min(upper.cols.x2, lower.cols.x2);
					if (overlapX2 < overlapX1) {
						return false; // нет реального перекрытия по X
					}

					// Вплотную по Y — считаем коридор «чистым»
					if (lower.rows.y1 - upper.rows.y2 <= 1) {
						return true;
					}

					// Проверяем вертикальный коридор между верхним и нижним регионами.
					// Сканируем только по тем колонкам (x), где области реально пересекаются.
					for (int x = overlapX1; x <= overlapX2; ++x) {
						for (int y = upper.rows.y2 + 1; y <= lower.rows.y1 - 1; ++y) {
							const char c = grid.At(y, x);
							if (c == ' ') {
								continue;
							}
							if (c == '=') {
								if (this->IsCellInAnyBar(bars, y, x)) {
									continue; // '=' внутри любого bar — это допустимо
								}
							}
							return false; // встретили «посторонний» символ → коридор «грязный»
						}
					}
					return true;
				}

				// Быстрый список всех bar-регионов (по узлам графа)
				std::vector<Model::Region> CollectBarRegions(const LayoutGraph& g) const {
					std::vector<Model::Region> bars{};
					for (const auto& n : g.Nodes()) {
						if (n.role == "bar") {
							bars.push_back(n.region);
						}
					}
					return bars;
				}

				// Принадлежит ли клетка (y,x) какому-либо bar-региону
				bool IsCellInAnyBar(const std::vector<Model::Region>& bars, int y, int x) const {
					for (const auto& r : bars) {
						if (r.rows.y1 <= y && y <= r.rows.y2 &&
							r.cols.x1 <= x && x <= r.cols.x2) {
							return true;
						}
					}
					return false;
				}

			private:
				double minRelativeXOverlap;
			};
		}
	}
}
