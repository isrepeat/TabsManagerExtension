#pragma once
#include <algorithm>
#include <string>
#include <vector>
#include <memory>

namespace AsciiMathParser {
	namespace Core {
		namespace Model {
			struct INode {
				virtual ~INode() = default;
			};

			struct NodesGroup {
				std::vector<std::unique_ptr<INode>> nodes;
			};

			struct Symbol : INode {
				std::string content;
			
				Symbol(std::string content)
					: content{ std::move(content) } {
				}
			};

			struct Frac : INode {
				NodesGroup num;
				NodesGroup den;

				Frac(
					NodesGroup num,
					NodesGroup den
				)
					: num{ std::move(num) }
					, den{ std::move(den) } {
				}
			};
		}
	}
}