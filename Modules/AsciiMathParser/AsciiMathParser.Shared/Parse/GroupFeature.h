#pragma once
#include <Helpers/Logger.h>

#include "../Model/Geometry.h"
#include "../Model/Group.h"
#include "IRegionFeature.h"

#include <stack>

namespace AsciiMathParser {
	namespace Core {
		namespace Parse {
            // Ищет на каждой строке пары фигурных скобок {…} с поддержкой вложенности (в пределах строки).
            // Для каждой пары создаёт Candidate:
            //   subregion = внутренность (x1+1..x2-1 @ y)
            //   structural skip = позиции самих скобок x1 и x2 (чтобы их не токенизировать)
            //   assembleFn => Group( parsed(inner) )
            class GroupFeature : public IRegionFeature {
            public:
                std::vector<Candidate> CollectChildren(
                    const Model::AsciiGrid& grid,
                    const Model::Region& region
                ) const override {
                    std::vector<Candidate> out{};
                    out.reserve(8);

                    for (int y = region.Top(); y <= region.Bottom(); ++y) {
                        std::vector<int> stackX{};
                        for (int x = region.Left(); x <= region.Right(); ++x) {
                            const char ch = grid.At(x, y);

                            if (ch == '{') {
                                stackX.push_back(x);
                                LOG_DEBUG_D("[Group] push  '{{' at y={} x={}", y, x);
                            }
                            else if (ch == '}') {
                                if (stackX.empty()) {
                                    LOG_DEBUG_D("[Group] stray '}}' at y={} x={} (ignored)", y, x);
                                    continue;
                                }

                                const int L = stackX.back();
                                stackX.pop_back();
                                const int R = x;

                                if (R <= L + 1) {
                                    LOG_DEBUG_D("[Group] empty '{{}}' pair at y={} L={} R={} (ignored)", y, L, R);
                                    continue;
                                }

                                // bbox всей группы — от '{' до '}'
                                const Model::Region bbox{
                                    Model::SpanX{ L, R },
                                    Model::SpanY{ y, y }
                                };

                                // внутренняя область — содержимое между скобками
                                const Model::Region inner{
                                    Model::SpanX{ L + 1, R - 1 },
                                    Model::SpanY{ y, y }
                                };

                                // Диагностика содержимого inner
                                std::string innerTxt;
                                for (int xx = inner.Left(); xx <= inner.Right(); ++xx) {
                                    innerTxt.push_back(grid.At(xx, y));
                                }

                                LOG_DEBUG_D(
                                    "[Group] pair y={} L={} R={} -> inner=[{}..{}] text='{}'",
                                    y, L, R, inner.Left(), inner.Right(), innerTxt
                                );

                                Candidate c{};
                                c.bbox = bbox;
                                c.subregions.push_back(inner);

                                // структурные пропуски: скрываем только сами скобки
                                c.mapRowToSkipRangesStructural[y].push_back(Model::SpanX{ L, L });
                                c.mapRowToSkipRangesStructural[y].push_back(Model::SpanX{ R, R });

                                c.assembleFn =
                                    [c, inner](std::vector<std::vector<std::ex::unique_ptr<Model::INode>>>&& subtrees)
                                    -> std::ex::unique_ptr<Model::INode> {
                                    // внутренняя группа — то, что распарсили в единственном подрегионе
                                    Model::NodesGroup innerGroup{ std::move(subtrees.at(0)) };

                                    // bbox группы можно оставить как c.bbox, либо объединить с inner (на случай расширений)
                                    const Model::Region unionBox = Model::Geometry::UnionRegion(
                                        c.bbox,
                                        Model::Geometry::UnionRegionsOfNodes(innerGroup.nodes)
                                    );

                                    return std::ex::make_unique_ex<Model::Group>(
                                        std::move(innerGroup),
                                        unionBox
                                    );
                                    };

                                out.push_back(std::move(c));
                            }
                        }
                    }

                    // Упорядочим кандидатов (стабильно слева-направо, сверху-вниз) — просто для повторяемости
                    std::stable_sort(
                        out.begin(),
                        out.end(),
                        [](const Candidate& a, const Candidate& b) {
                            if (a.bbox.Left() != b.bbox.Left()) {
                                return a.bbox.Left() < b.bbox.Left();
                            }
                            return a.bbox.Top() < b.bbox.Top();
                        }
                    );

                    return out;
                }
            };
		}
	}
}