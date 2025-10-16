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
		}
	}
}