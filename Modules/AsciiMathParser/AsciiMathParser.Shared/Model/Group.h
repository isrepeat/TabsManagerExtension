#pragma once
#include "Geometry.h"
#include "INode.h"

#include <memory>
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		namespace Model {
            class Group : public INode {
            public:
                Group(
                    NodesGroup inner,
                    Region bbox
                )
                    : inner{ std::move(inner) }
                    , region{ bbox } {
                }

                Region GetRegion() const override {
                    return this->region;
                }

                // Смотреть, но не трогать
                const NodesGroup& Inner() const {
                    return this->inner;
                }

                // Забрать владение узлами (move)
                NodesGroup ReleaseInner() {
                    return std::move(this->inner);
                }

            private:
                NodesGroup inner;
                Region region;
            };
		}
	}
}