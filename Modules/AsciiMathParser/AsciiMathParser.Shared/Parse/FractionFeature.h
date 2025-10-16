#pragma once
#include <Helpers/Logger.h>

#include "../Detect/FractionBarDetector.h"
#include "../Model/Fraction.h"
#include "../Model/Geometry.h"
#include "IRegionFeature.h"

namespace AsciiMathParser {
	namespace Core {
		namespace Parse {
			class FractionFeature final : public IRegionFeature {
			public:
				explicit FractionFeature(
					const std::vector<Detect::FractionBar>& bars
				)
					: bars{ bars } {
				}


				std::vector<Candidate> CollectChildren(
					const Model::AsciiGrid&,
					const Model::Region& region
				) const override {
					// 1) берём бары, полностью попавшие в наш region
					std::vector<std::size_t> inside{};
					inside.reserve(this->bars.size());
					for (std::size_t i = 0; i < this->bars.size(); ++i) {
						const auto& b = this->bars[i];
						const bool ok =
							region.ContainsY(b.barRegion.y) &&
							region.ContainsX(b.barRegion.Left()) &&
							region.ContainsX(b.barRegion.Right());
						if (ok) {
							inside.push_back(i);
						}
					}

					// 2) отфильтровать вложенные: top — те, чья линия бара не лежит внутри num/den других
					std::vector<std::size_t> topIdx{};
					topIdx.reserve(inside.size());
					for (auto i : inside) {
						const auto barReg = this->bars[i].barRegion.ToRegion();
						bool nested = false;

						for (auto j : inside) {
							if (i == j) {
								continue;
							}
							const auto& bj = this->bars[j];
							if (bj.numRegion.ContainsSubregion(barReg) ||
								bj.denRegion.ContainsSubregion(barReg)) {
								nested = true;
								break;
							}
						}

						if (!nested) {
							topIdx.push_back(i);
						}
					}

					// 3) собрать кандидатов (самодостаточные)
					std::vector<Candidate> out{};
					out.reserve(topIdx.size());

					for (auto idx : topIdx) {
						const auto& b = this->bars[idx];

						Candidate candidate{};

						// bbox = union(bar ∪ num ∪ den)
						candidate.bbox = Model::Geometry::UnionRegions(
							b.barRegion.ToRegion(),
							b.numRegion,
							b.denRegion
						);

						// подрегионы для рекурсии
						candidate.subregions = {
							b.numRegion,
							b.denRegion
						};

						// структурные пропуски: ТОЛЬКО линия бара
						// (RegionWalker потом сам склеит их по всем top-кандидатам)
						if (region.ContainsY(b.barRegion.y)) {
							const int L = std::max(region.Left(), b.barRegion.Left());
							const int R = std::min(region.Right(), b.barRegion.Right());
							if (L <= R) {
								candidate.mapRowToSkipRangesStructural[b.barRegion.y].push_back(Model::SpanX{ L, R });
							}
						}

						// assemble-замыкание: строит Fraction из готовых subtrees
						candidate.assembleFn =
							[b](std::vector<std::vector<std::unique_ptr<Model::INode>>>&& subtrees)
							-> std::unique_ptr<Model::INode> {
							auto numGroup = Model::NodesGroup{ std::move(subtrees.at(0)) };
							auto denGroup = Model::NodesGroup{ std::move(subtrees.at(1)) };

							const Model::Region unionRegion =
								Model::Geometry::UnionRegions(
									b.barRegion.ToRegion(),
									b.numRegion,
									b.denRegion
								);

							return std::make_unique<Model::Fraction>(
								std::move(numGroup),
								std::move(denGroup),
								unionRegion
							);
							};

						out.push_back(std::move(candidate));
					}

					// читаемость дампа: слева→направо, сверху→вниз
					std::sort(
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

			private:
				std::vector<Detect::FractionBar> bars;
			};
		}
	}
}