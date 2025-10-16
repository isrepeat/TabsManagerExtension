#pragma once
#include "../Detect/FractionBarDetector.h"
#include "../Model/Fraction.h"
#include "../Model/Geometry.h"
#include "IRegionFeature.h"

namespace AsciiMathParser {
	namespace Core {
		namespace Parse {
			class FractionFeature : public IRegionFeature {
			public:
				// Можно принимать бары извне (кэш сканирования по всей сетке),
				// либо.detectить внутри CollectChildren по grid — на ваше усмотрение.
				FractionFeature(std::vector<Detect::FractionBar> bars)
					: bars_{ std::move(bars) } {
				}

				std::vector<FeatureChild> CollectChildren(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) override {
					this->cache_.clear();

					std::vector<Detect::FractionBar> inside{};
					for (const auto& b : this->bars_) {
						if (FractionFeature::IsBarInsideRegion(b, region)) {
							inside.push_back(b);
						}
					}

					// Оставить только top-level внутри данного region
					std::vector<Detect::FractionBar> top{};
					for (std::size_t i = 0; i < inside.size(); ++i) {
						bool nested = false;
						for (std::size_t j = 0; j < inside.size(); ++j) {
							if (i == j) {
								continue;
							}
							if (FractionFeature::IsBarInsideRegion(inside[i], inside[j].numRegion) ||
								FractionFeature::IsBarInsideRegion(inside[i], inside[j].denRegion)) {
								nested = true;
								break;
							}
						}
						if (!nested) {
							top.push_back(inside[i]);
						}
					}

					// Преобразовать в FeatureChild
					std::vector<FeatureChild> out{};
					out.reserve(top.size());
					for (const auto& b : top) {
						FeatureChild ch{
							.owner = this,
							.bbox = Model::Geometry::UnionRegions(
								b.barRegion.ToRegion(),
								b.numRegion,
								b.denRegion
							),
							.ownedSubregions = { b.numRegion, b.denRegion },
							.implData = nullptr
						};
						// Храним копию бара рядом, если нужно — можно повесить mini-storage/индекс
						this->cache_.push_back(b);
						out.push_back(ch);
					}

					// Стабильный порядок чтения
					std::sort(
						out.begin(),
						out.end(),
						[](const FeatureChild& a, const FeatureChild& b) {
							if (a.bbox.Left() != b.bbox.Left()) {
								return a.bbox.Left() < b.bbox.Left();
							}
							return a.bbox.Top() < b.bbox.Top();
						}
					);
					return out;
				}

				void AppendSkipSpans(
					const Model::Region& currentRegion,
					const FeatureChild& child,
					std::unordered_map<int, std::vector<Model::SpanX>>& skipByRow
				) const override {
					// 1) Полоса бара в пределах currentRegion.cols
					for (const auto& b : this->cache_) {
						if (!currentRegion.ContainsY(b.barRegion.y)) {
							continue;
						}
						const int L = std::max(currentRegion.Left(), b.barRegion.Left());
						const int R = std::min(currentRegion.Right(), b.barRegion.Right());
						if (L <= R) {
							skipByRow[b.barRegion.y].push_back(Model::SpanX{ L, R });
						}
					}

					// 2) Окна num/den — клип по currentRegion
					for (const auto& b : this->cache_) {
						for (int y = b.numRegion.Top(); y <= b.numRegion.Bottom(); ++y) {
							if (!currentRegion.ContainsY(y)) {
								continue;
							}
							const int L = std::max(currentRegion.Left(), b.numRegion.Left());
							const int R = std::min(currentRegion.Right(), b.numRegion.Right());
							if (L <= R) {
								skipByRow[y].push_back(Model::SpanX{ L, R });
							}
						}
						for (int y = b.denRegion.Top(); y <= b.denRegion.Bottom(); ++y) {
							if (!currentRegion.ContainsY(y)) {
								continue;
							}
							const int L = std::max(currentRegion.Left(), b.denRegion.Left());
							const int R = std::min(currentRegion.Right(), b.denRegion.Right());
							if (L <= R) {
								skipByRow[y].push_back(Model::SpanX{ L, R });
							}
						}
					}
				}

				std::unique_ptr<Model::INode> BuildNode(
					const Model::AsciiGrid& grid,
					const Model::Region& currentRegion,
					const FeatureChild& child,
					const IRegionWalker& walker
				) const override {
					// Находим соответствующий bar для child.bbox (в реальном коде храните индекс)
					const Detect::FractionBar* bar = nullptr;
					for (const auto& b : this->cache_) {
						Model::Region bbox = Model::Geometry::UnionRegions(
							b.barRegion.ToRegion(),
							b.numRegion,
							b.denRegion
						);
						if (bbox.Left() == child.bbox.Left() &&
							bbox.Right() == child.bbox.Right() &&
							bbox.Top() == child.bbox.Top() &&
							bbox.Bottom() == child.bbox.Bottom()
							) {
							bar = &b;
							break;
						}
					}

					// Рекурсивно собрать детей num/den
					auto numNodes = walker.ParseRegion(
						grid,
						bar->numRegion
					);
					auto denNodes = walker.ParseRegion(
						grid,
						bar->denRegion
					);

					// Собрать Frac узел (как сейчас в RegionParser)
					return std::make_unique<Model::Fraction>(
						Model::NodesGroup{ std::move(numNodes) },
						Model::NodesGroup{ std::move(denNodes) },
						bar->barRegion.ToRegion()
					);
				}

			private:
				static bool IsBarInsideRegion(
					const Detect::FractionBar& bar,
					const Model::Region& region
				) {
					if (!region.ContainsY(bar.barRegion.y)) {
						return false;
					}
					if (!region.ContainsX(bar.barRegion.Left())) {
						return false;
					}
					if (!region.ContainsX(bar.barRegion.Right())) {
						return false;
					}
					return true;
				}

			private:
				std::vector<Detect::FractionBar> bars_;
				// Локальный кэш «видимых» баров для текущей CollectChildren; в проде — лучше индексировать
				mutable std::vector<Detect::FractionBar> cache_;
			};
		}
	}
}