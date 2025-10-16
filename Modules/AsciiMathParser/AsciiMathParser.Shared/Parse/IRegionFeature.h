#pragma once
#include "../Model/INode.h"
#include "../Model/Grid.h"
#include "IRegionWalker.h"

#include <unordered_map>
#include <vector>
#include <memory>

namespace AsciiMathParser {
	namespace Core {
		namespace Parse {
			using FeatureChildId = std::uintptr_t; // фича сама решает, что класть (обычно индекс)

			struct PlannedChild {
				const class IRegionFeature* owner;   // кто породил
				Model::Region               bbox;    // общий бокс ребёнка
				std::vector<Model::Region>  subregions; // куда рекурсить
				FeatureChildId              id;      // feature-local ID (напр., индекс бара)
			};


			class IRegionFeature {
			public:
				virtual ~IRegionFeature() = default;

				// Найти детей ТОЛЬКО внутри 'region' (без рекурсии)
				virtual std::vector<PlannedChild> CollectChildren(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) const = 0;

				// Добавить пропуски только под КОНКРЕТНОГО ребёнка (полоса бара, окна и т.д.)
				virtual void AppendSkips(
					const Model::Region& currentRegion,
					FeatureChildId id,
					std::unordered_map<int, std::vector<Model::SpanX>>& skipByRow
				) const = 0;

				// Сконструировать финальный узел, получив уже распарсенные поддеревья
				// subtrees.size() == subregions.size() из PlannedChild.
				virtual std::unique_ptr<Model::INode> Assemble(
					FeatureChildId id,
					std::vector<std::vector<std::unique_ptr<Model::INode>>>&& subtrees
				) const = 0;
			};
		}
	}
}