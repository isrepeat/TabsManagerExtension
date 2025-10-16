#pragma once
#include "../Model/INode.h"
#include "../Model/Grid.h"
#include <memory>
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		namespace Parse {
			class IRegionWalker {
			public:
				virtual ~IRegionWalker() = default;

				virtual std::vector<std::unique_ptr<Model::INode>> ParseRegion(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) const = 0;
			};
		}
	}
}