#pragma once
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
				std::vector<std::unique_ptr<INode>> nodes;
			};

			//
			// Symbol
			//
			class Symbol : public INode {
			public:
				Symbol(
					std::string content,
					int x,
					int y
				)
					: content{ std::move(content) }
					, region{
						SpanX{ x, x },
						SpanY{ y, y }
					} {
				}

				Region GetRegion() const override {
					return this->region;
				}

				const std::string& Content() const {
					return this->content;
				}

			private:
				const std::string content;
				const Region region;
			};


			//
			// Frac
			//
			class Frac : public INode {
			public:
				Frac(
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