#pragma once
#include "Geometry.h"
#include "INode.h"

#include <optional>
#include <memory>

namespace AsciiMathParser {
	namespace Core {
		namespace Model {
            // Объединённый узел для x^{sup}_{sub} / x^{sup} / x_{sub}
            class Script : public INode {
            public:
                Script(
                    std::ex::unique_ptr<INode> base,
                    std::optional<NodesGroup> sup,
                    std::optional<NodesGroup> sub
                )
                    : base{ std::move(base) }
                    , sup{ std::move(sup) }
                    , sub{ std::move(sub) }
                    , region{ this->ComputeRegion_() } {
                }

                Region GetRegion() const override {
                    return this->region;
                }

                const INode& Base() const {
                    return *(this->base);
                }

                const std::optional<NodesGroup>& Sup() const {
                    return this->sup;
                }

                const std::optional<NodesGroup>& Sub() const {
                    return this->sub;
                }

            private:
                Region ComputeRegion_() const {
                    std::optional<Region> acc{ this->base->GetRegion() };

                    if (this->sup.has_value()) {
                        const Region rs = Model::Geometry::UnionRegionsOfNodes(this->sup->nodes);
                        acc.emplace(Model::Geometry::UnionRegion(*(acc), rs));
                    }
                    if (this->sub.has_value()) {
                        const Region rb = Model::Geometry::UnionRegionsOfNodes(this->sub->nodes);
                        acc.emplace(Model::Geometry::UnionRegion(*(acc), rb));
                    }
                    return acc.value();
                }

            private:
                std::ex::unique_ptr<INode> base;
                std::optional<NodesGroup> sup;
                std::optional<NodesGroup> sub;
                Region region;
            };
		}
	}
}