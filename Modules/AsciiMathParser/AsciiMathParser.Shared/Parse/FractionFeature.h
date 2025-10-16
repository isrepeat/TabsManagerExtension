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
					: bars_{ bars } {
				}


				std::vector<Candidate> CollectChildren(
					const Model::AsciiGrid&,
					const Model::Region& region
				) const override {
					// внутри региона → top-level по «бар не внутри num/den другого»
					std::vector<std::size_t> inside{};
					for (std::size_t i = 0; i < this->bars_.size(); ++i) {
						const auto& b = this->bars_[i];
						if (FractionFeature::IsBarInside(b, region)) {
							inside.push_back(i);
						}
					}

					LOG_DEBUG_D(
						"[FractionFeature] inside bars: {}",
						(int)inside.size()
					);

					std::vector<std::size_t> topIdx{};
					for (auto i : inside) {
						bool nested = false;
						for (auto j : inside) {
							if (i == j) {
								continue;
							}
							const auto& bj = this->bars_[j];
							const auto barReg = this->bars_[i].barRegion.ToRegion();
							if (bj.numRegion.ContainsSubregion(barReg) ||
								bj.denRegion.ContainsSubregion(barReg)
								) {
								nested = true;
								break;
							}
						}
						if (!nested) {
							topIdx.push_back(i);
						}
					}

					LOG_DEBUG_D(
						"[FractionFeature] top bars: {}",
						(int)topIdx.size()
					);

					std::vector<Candidate> out{};
					out.reserve(topIdx.size());

					for (auto idx : topIdx) {
						const auto& b = this->bars_[idx];
						LOG_DEBUG_D(
							"  top bar id={} bar=[{}..{} @y{}], num=[{}..{} x {}..{}], den=[{}..{} x {}..{}]",
							(unsigned long long)idx,
							b.barRegion.Left(), b.barRegion.Right(), b.barRegion.y,
							b.numRegion.Left(), b.numRegion.Right(), b.numRegion.Top(), b.numRegion.Bottom(),
							b.denRegion.Left(), b.denRegion.Right(), b.denRegion.Top(), b.denRegion.Bottom()
						);

						Candidate ch{
							.owner = this,
							.id = static_cast<int>(idx),
							.bbox = Model::Geometry::UnionRegions(
								b.barRegion.ToRegion(),
								b.numRegion,
								b.denRegion
							),
							.subregions = { b.numRegion, b.denRegion }
						};

						out.push_back(std::move(ch));
					}
					std::sort(
						out.begin(), out.end(),
						[](const Candidate& a, const Candidate& b) {
							if (a.bbox.Left() != b.bbox.Left()) {
								return a.bbox.Left() < b.bbox.Left();
							}
							return a.bbox.Top() < b.bbox.Top();
						}
					);
					return out;
				}


				// В skip добавляем ТОЛЬКО линию бара. Num/Den не трогаем.
				void AppendSkips(
					const Model::Region& regionCurrent,
					int candidateId,
					std::unordered_map<int, std::vector<Model::SpanX>>& skipMap
				) const override {
					const auto idx = static_cast<size_t>(candidateId);
					const auto& b = this->bars_.at(idx);

					const int y = b.barRegion.y;
					if (!regionCurrent.ContainsY(y)) {
						return;
					}

					const int left = std::max(regionCurrent.Left(), b.barRegion.Left());
					const int right = std::min(regionCurrent.Right(), b.barRegion.Right());

					if (left <= right) {
						skipMap[y].push_back(Model::SpanX{ left, right });
						LOG_DEBUG_D("  skip bar y={} [{}..{}]", y, left, right);
					}
				}

				// Собираем Fraction-узел: регион узла = union(bar ∪ num ∪ den).
				std::unique_ptr<Model::INode> Assemble(
					int candidateId,
					std::vector<std::vector<std::unique_ptr<Model::INode>>>&& subtrees
				) const override {
					const auto idx = static_cast<size_t>(candidateId);
					const auto& b = this->bars_.at(idx);

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
				}

			private:
				static bool IsBarInside(
					const Detect::FractionBar& bar,
					const Model::Region& region
				) {
					return region.ContainsY(bar.barRegion.y) &&
						region.ContainsX(bar.barRegion.Left()) &&
						region.ContainsX(bar.barRegion.Right());
				}

			private:
				const std::vector<Detect::FractionBar> bars_;
			};
		}
	}
}