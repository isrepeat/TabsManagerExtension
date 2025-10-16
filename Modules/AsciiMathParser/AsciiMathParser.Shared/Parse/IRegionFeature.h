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
			struct FeatureChild {
				// Владелец — та фича, что породила этого ребёнка
				class IRegionFeature* owner;

				// Глобальный bbox узла (например, бар + num/den + tight-обвязка)
				Model::Region bbox;

				// Набор «подвластных» под-регионов (куда провалится рекурсия):
				// у дроби это numRegion и denRegion; у радикала — radicand и, возможно, index.
				std::vector<Model::Region> ownedSubregions;

				// Доп. произвольные данные фичи (type-erased handle)
				void* implData;
			};


			class IRegionFeature {
			public:
				virtual ~IRegionFeature() = default;

				// Найти детей в пределах region (без рекурсии)
				virtual std::vector<FeatureChild> CollectChildren(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) = 0;

				// Добавить вклад в skip-карту текущего уровня для данного ребёнка
				virtual void AppendSkipSpans(
					const Model::Region& currentRegion,
					const FeatureChild& child,
					std::unordered_map<int, std::vector<Model::SpanX>>& skipByRow
				) const = 0;

				// Построить финальный INode для ребёнка; рекурсивно разобрать его под-регионы через walker
				virtual std::unique_ptr<Model::INode> BuildNode(
					const Model::AsciiGrid& grid,
					const Model::Region& currentRegion,
					const FeatureChild& child,
					const IRegionWalker& walker
				) const = 0;
			};
		}
	}
}