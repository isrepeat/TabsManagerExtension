#pragma once
#include "../Model/Geometry.h"
#include "../Model/Symbol.h"
#include "IRegionFeature.h"
#include "IRegionWalker.h"

#include <unordered_map>
#include <algorithm>
#include <optional>

namespace AsciiMathParser {
	namespace Core {
		namespace Parse {
			class RegionWalker : public IRegionWalker {
			public:
				explicit RegionWalker(std::vector<std::unique_ptr<IRegionFeature>>&& features)
					: features_{ std::move(features) } {
				}

				std::vector<std::unique_ptr<Model::INode>> ParseRegion(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) const override {
					// 1) Сбор детей у всех фич
					std::vector<FeatureChild> allChildren{};
					for (const auto& featurePtr : this->features_) {
						std::vector<FeatureChild> locals = featurePtr->CollectChildren(
							grid,
							region
						);
						for (auto& ch : locals) {
							allChildren.push_back(std::move(ch));
						}
					}

					// 2) Оставить только верхнеуровневых
					std::vector<FeatureChild> top = this->FilterTopLevel(allChildren);

					// 3) Рекурсивно построить узлы детей
					std::vector<std::unique_ptr<Model::INode>> childNodes{};
					for (const auto& ch : top) {
						// Каждая фича сама знает, как рекурсивно обойти свои под-регионы:
						// ей передаётся ссылка на «walker» (this) для вложенного парсинга.
						childNodes.push_back(
							ch.owner->BuildNode(
								grid,
								region,
								ch,
								*this
							)
						);
					}

					// 4) Построить skip-карту из вкладов всех верхнеуровневых детей
					std::unordered_map<int, std::vector<Model::SpanX>> skipByRow{};
					for (const auto& ch : top) {
						ch.owner->AppendSkipSpans(
							region,
							ch,
							skipByRow
						);
					}
					// Слить интервалы на строках
					for (auto& kv : skipByRow) {
						RegionWalker::MergeRowSpans(kv.second);
					}

					// 5) Токенизировать оставшееся в Symbol
					auto symNodes = this->TokenizeToSymbols(
						grid,
						region,
						skipByRow
					);

					// 6) Слить по reading-order
					auto merged = this->MergeByReadingOrder(
						std::move(childNodes),
						std::move(symNodes)
					);
					return merged;
				}

			private:
				std::vector<std::unique_ptr<Model::INode>> TokenizeToSymbols(
					const Model::AsciiGrid& grid,
					const Model::Region& region,
					const std::unordered_map<int, std::vector<Model::SpanX>>& skipByRow
				) const {
					std::vector<std::unique_ptr<Model::INode>> out{};

					// Проходим все строки текущего региона сверху вниз.
					for (int y = region.Top(); y <= region.Bottom(); ++y) {

						// Получаем список X-интервалов, которые нужно пропустить на этой строке.
						// Если строка не имеет пропусков — используем пустой список.
						const auto it = skipByRow.find(y);
						static const std::vector<Model::SpanX> kEmpty{};
						const auto& rowSkips = (it == skipByRow.end() ? kEmpty : it->second);

						// Получаем «разрешённые» промежутки X в этой строке,
						// то есть области, не попадающие в skipRow.
						for (const auto& allowedRange : RegionWalker::AllowedRanges(region, rowSkips)) {
							// Начинаем обход разрешённого диапазона слева направо.
							int x = allowedRange.Left();

							// Идём по всем символам до конца разрешённого диапазона.
							while (x <= allowedRange.Right()) {
								// Пропускаем все пробелы подряд.
								while (x <= allowedRange.Right() && grid.At(x, y) == ' ') {
									++x;
								}

								// Если после пропусков дошли до конца диапазона — строка исчерпана.
								if (x > allowedRange.Right()) {
									break;
								}

								// Здесь x указывает на первый непробельный символ токена.
								const int startX = x;
								std::string token{};

								// Собираем последовательность непробельных символов до конца токена
								// (пока не встретим пробел или конец диапазона).
								while (x <= allowedRange.Right() && grid.At(x, y) != ' ') {
									token.push_back(grid.At(x, y));
									++x;
								}

								// Если собран непустой токен — создаём узел Symbol и сохраняем его.
								if (!token.empty()) {
									out.push_back(std::make_unique<Model::Symbol>(
										std::move(token),
										startX, // X-позиция первого символа токена
										y       // строка, на которой токен расположен
									));
								}
							}
						}
					}

					return out;
				}

				std::vector<std::unique_ptr<Model::INode>> MergeByReadingOrder(
					std::vector<std::unique_ptr<Model::INode>> fracNodes,
					std::vector<std::unique_ptr<Model::INode>> symNodes
				) const {
					std::vector<std::unique_ptr<Model::INode>> all{};
					all.reserve(fracNodes.size() + symNodes.size());

					for (auto& p : fracNodes) {
						all.push_back(std::move(p));
					}
					for (auto& p : symNodes) {
						all.push_back(std::move(p));
					}

					std::sort(
						all.begin(),
						all.end(),
						[](const std::unique_ptr<Model::INode>& a, const std::unique_ptr<Model::INode>& b) {
							const auto ra = a->GetRegion();
							const auto rb = b->GetRegion();
							if (ra.Left() != rb.Left()) {
								return ra.Left() < rb.Left(); // первичный ключ — левый край
							}
							return ra.Top() < rb.Top(); // вторичный — верх
						}
					);

					return all;
				}


				std::vector<FeatureChild> FilterTopLevel(
					const std::vector<FeatureChild>& input
				) const {
					// Правило: если bbox A целиком попадает в любой ownedSubregion B — A не top-level.
					std::vector<FeatureChild> out{};
					for (std::size_t i = 0; i < input.size(); ++i) {
						bool nested = false;
						for (std::size_t j = 0; j < input.size(); ++j) {
							if (i == j) {
								continue;
							}
							for (const auto& owned : input[j].ownedSubregions) {
								if (RegionWalker::ContainsStrict(owned, input[i].bbox)) {
									nested = true;
									break;
								}
							}
							if (nested) {
								break;
							}
						}
						if (!nested) {
							out.push_back(input[i]);
						}
					}
					// Стабильный порядок: x asc, потом y asc — как и было
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

				// Разность по X: всё, что можно читать в строке y внутри region (без skip-интервалов).
				static std::vector<Model::SpanX> AllowedRanges(
					const Model::Region& region,
					const std::vector<Model::SpanX>& skipRow
				) {
					std::vector<Model::SpanX> free{};

					if (skipRow.empty()) {
						free.push_back(Model::SpanX{ region.Left(), region.Right() });
						return free;
					}

					int cur = region.Left();
					for (const auto& s : skipRow) {
						if (s.Left() > cur) {
							free.push_back(Model::SpanX{ cur, s.Left() - 1 });
						}
						cur = std::max(cur, s.Right() + 1);
					}

					if (cur <= region.Right()) {
						free.push_back(Model::SpanX{ cur, region.Right() });
					}

					return free;
				}


				// Сливает перекрывающиеся или примыкающие интервалы на одной строке:
				// [3..8], [7..10] → [3..10]. Это делает карту skip компактной и
				// уменьшает количество прыжков токенайзера.
				static void MergeRowSpans(std::vector<Model::SpanX>& spans) {
					std::sort(
						spans.begin(),
						spans.end(),
						[](const Model::SpanX& a, const Model::SpanX& b) {
							if (a.Left() != b.Left()) {
								return a.Left() < b.Left();
							}
							return a.Right() < b.Right();
						}
					);

					std::vector<Model::SpanX> merged{};
					std::optional<std::pair<int, int>> cur{};

					for (const auto& s : spans) {
						if (!cur.has_value()) {
							cur = std::make_pair(s.Left(), s.Right());
						}
						else {
							auto& [cx1, cx2] = *cur;
							if (s.Left() <= cx2 + 1) {
								if (s.Right() > cx2) cx2 = s.Right();
							}
							else {
								merged.push_back(Model::SpanX{ cx1, cx2 });
								cur = std::make_pair(s.Left(), s.Right());
							}
						}
					}
					if (cur.has_value()) {
						merged.push_back(Model::SpanX{ cur->first, cur->second });
					}
					spans = std::move(merged);
				}

				static bool ContainsStrict(
					const Model::Region& bigRegion,
					const Model::Region& smallRegion
				) {
					if (smallRegion.Left() < bigRegion.Left()) {
						return false;
					}
					if (smallRegion.Right() > bigRegion.Right()) {
						return false;
					}
					if (smallRegion.Top() < bigRegion.Top()) {
						return false;
					}
					if (smallRegion.Bottom() > bigRegion.Bottom()) {
						return false;
					}
					return true;
				}

			private:
				std::vector<std::unique_ptr<IRegionFeature>> features_;
			};
		}
	}
}