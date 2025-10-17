#pragma once
#include <Helpers/Std/Extensions/memoryEx.h>

#include "Geometry.h"
#include "Grid.h"

#include <algorithm>
#include <string>
#include <vector>
#include <memory>

namespace AsciiMathParser {
	namespace Core {
		namespace Model {
			class INode {
			public:
				virtual ~INode() = default;
				virtual Region GetRegion() const = 0;
			};


			struct NodesGroup {
				std::vector<std::ex::unique_ptr<INode>> nodes;
			};
		}
	}
}