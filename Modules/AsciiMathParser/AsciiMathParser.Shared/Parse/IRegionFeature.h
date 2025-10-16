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
			//using FeatureChildId = std::uintptr_t; // фича сама решает, что класть (обычно индекс)

			struct IRegionFeature;

			// Один найденный «кандидат» фичи (дробь, корень и т.п.) внутри region.
			struct Candidate final {
				const IRegionFeature* owner = nullptr;  // владелец-детектор
				int id = 0;								// индекс внутри детектора (например bars_[id])
				Model::Region bbox;						// общий bouding-box
				std::vector<Model::Region> subregions;  // области для рекурсии (Num/Den/...)
			};


			// Интерфейс любого детектора регионов.
			struct IRegionFeature {
				virtual ~IRegionFeature() {}

				// Найти своих кандидатов в пределах region.
				virtual std::vector<Candidate> CollectChildren(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) const = 0;

				// Добавить «структурные» зоны в skip (например, линию бара).
				// ВНИМАНИЕ: содержимое (Num/Den/...) не добавлять в skip — оно парсится рекурсией.
				virtual void AppendSkips(
					const Model::Region& regionCurrent,
					int candidateId,
					std::unordered_map<int, std::vector<Model::SpanX>>& skipMap
				) const = 0;

				// Собрать финальный узел из уже разобранных поддеревьев.
				virtual std::unique_ptr<Model::INode> Assemble(
					int candidateId,
					std::vector<std::vector<std::unique_ptr<Model::INode>>>&& subtrees
				) const = 0;
			};
		}
	}
}