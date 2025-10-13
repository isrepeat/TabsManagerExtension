#pragma once
#include "../Detect/IFeatureDetector.h"
#include "../Model/Grid.h"
#include "LayoutGraph.h"
#include <unordered_map>
#include <string_view>
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		namespace Layout {
			struct RelationBuilderContext {
				LayoutGraph& graph;
				const Model::AsciiGrid& grid;

				// Индексы узлов по ролям: "bar", "sqrt", "bracket", "text-run" и т.д.
				std::unordered_map<std::string_view, std::vector<const LayoutNode*>> byRole;

				RelationBuilderContext(
					LayoutGraph& g,
					const Model::AsciiGrid& grd
				)
					: graph{ g }
					, grid{ grd } {
					for (auto& node : g.Nodes()) {
						this->byRole[node.role].push_back(&node);
					}
				}
			};
		}
	}
}
