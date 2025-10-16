#pragma once
#include "Geometry.h"
#include "INode.h"

#include <algorithm>
#include <string>
#include <vector>
#include <memory>

namespace AsciiMathParser {
	namespace Core {
		namespace Model {
			class Fraction : public INode {
			public:
				Fraction(
					NodesGroup num,
					NodesGroup den,
					Region barRegion
				)
					: num{ std::move(num) }
					, den{ std::move(den) }
					, region{ Geometry::UnionRegions(
						barRegion,
						Geometry::UnionRegionsOfNodes(this->num.nodes),
						Geometry::UnionRegionsOfNodes(this->den.nodes)
					) } {
				}

				Region GetRegion() const override {
					return this->region;
				}

				const NodesGroup& Num() const {
					return this->num;
				}

				const NodesGroup& Den() const {
					return this->den;
				}

			private:
				const NodesGroup num;
				const NodesGroup den;
				const Region region;
			};
		}
	}
}