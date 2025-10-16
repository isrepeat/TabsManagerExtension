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


				std::vector<PlannedChild> CollectChildren(
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

					std::vector<PlannedChild> out{};
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

						PlannedChild ch{
							.owner = this,
							.bbox = Model::Geometry::UnionRegions(
								b.barRegion.ToRegion(),
								b.numRegion,
								b.denRegion
							),
							.subregions = { b.numRegion, b.denRegion },
							.id = static_cast<FeatureChildId>(idx)
						};

						out.push_back(std::move(ch));
					}
					std::sort(
						out.begin(), out.end(),
						[](const PlannedChild& a, const PlannedChild& b) {
							if (a.bbox.Left() != b.bbox.Left()) {
								return a.bbox.Left() < b.bbox.Left();
							}
							return a.bbox.Top() < b.bbox.Top();
						}
					);
					return out;
				}


				void AppendSkips(
					const Model::Region& cur,
					FeatureChildId id,
					std::unordered_map<int, std::vector<Model::SpanX>>& skip
				) const override {
					const auto& b = this->bars_[static_cast<std::size_t>(id)];

					auto clipRegion = [&](
						const Model::Region& r,
						const char* tag
						) {
							for (int y = r.Top(); y <= r.Bottom(); ++y) {
								if (!cur.ContainsY(y)) {
									continue;
								}
								const int L = std::max(cur.Left(), r.Left());
								const int R = std::min(cur.Right(), r.Right());
								if (L <= R) {
									skip[y].push_back(Model::SpanX{ L, R });
									LOG_DEBUG_D(
										"  skip {} y={} [{}..{}]",
										tag,
										y,
										L,
										R
									);
								}
							}
						};

					// полоса бара
					if (cur.ContainsY(b.barRegion.y)) {
						const int L = std::max(cur.Left(), b.barRegion.Left());
						const int R = std::min(cur.Right(), b.barRegion.Right());
						if (L <= R) {
							skip[b.barRegion.y].push_back(Model::SpanX{ L, R });
							LOG_DEBUG_D(
								"  skip bar y={} [{}..{}]",
								b.barRegion.y,
								L,
								R
							);
						}
					}

					clipRegion(
						b.numRegion,
						"num"
					);
					clipRegion(
						b.denRegion,
						"den"
					);
				}


				std::unique_ptr<Model::INode> Assemble(
					FeatureChildId id,
					std::vector<std::vector<std::unique_ptr<Model::INode>>>&& subtrees
				) const override {
					const auto& b = this->bars_[static_cast<std::size_t>(id)];

					LOG_DEBUG_D(
						"[FractionFeature] Assemble id={} (Num={}, Den={})",
						(unsigned long long)id,
						(int)subtrees[0].size(),
						(int)subtrees[1].size()
					);

					return std::make_unique<Model::Fraction>(
						Model::NodesGroup{ std::move(subtrees[0]) },
						Model::NodesGroup{ std::move(subtrees[1]) },
						b.barRegion.ToRegion()
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
				//const std::vector<Detect::FractionBar>& bars_; // только ссылка, без кешей
				const std::vector<Detect::FractionBar> bars_; // только ссылка, без кешей
			};
		}
	}
}