#pragma once
#include <Helpers/Logger.h>

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
			class RegionWalker {
			public:
				explicit RegionWalker(
					std::vector<std::unique_ptr<IRegionFeature>>&& features
				)
					: features_{ std::move(features) } {
				}

				// Внешний вход: без наследованных skip’ов
				std::vector<std::unique_ptr<Model::INode>> ParseRegion(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) const {
					std::unordered_map<int, std::vector<Model::SpanX>> empty{};
					return this->ParseRegion(grid, region, empty);
				}

				// Основной вариант: с наследованными skip’ами
				std::vector<std::unique_ptr<Model::INode>> ParseRegion(
					const Model::AsciiGrid& grid,
					const Model::Region& region,
					const std::unordered_map<int, std::vector<Model::SpanX>>& inheritedSkip
				) const {
					LOG_FUNCTION_SCOPE("@@@ ParseRegion(): bbox=[x:{}..{} y:{}..{}]",
						region.Left(), region.Right(), region.Top(), region.Bottom());

					const auto top = this->CollectTopLevel_(grid, region);
					auto skip = this->BuildSkip_(region, top, inheritedSkip);

					// Собираем рекурсивно фичи
					std::vector<std::unique_ptr<Model::INode>> feats{};
					feats.reserve(top.size());

					for (const auto& ch : top) {
						std::vector<std::vector<std::unique_ptr<Model::INode>>> subtrees{};
						subtrees.reserve(ch.subregions.size());

						for (const auto& sub : ch.subregions) {
							auto subNodes = this->ParseRegion(grid, sub, skip);
							subtrees.push_back(std::move(subNodes));
						}

						feats.push_back(ch.owner->Assemble(ch.id, std::move(subtrees)));
					}

					this->OccupyFeatureRegions_(region, feats, skip);

					auto symbols = this->TokenizeToSymbols(grid, region, skip);
					auto merged = this->MergeByReadingOrder(std::move(feats), std::move(symbols));

					return merged;
				}

			private:
				// 1) Собрать top-level кандидатов
				std::vector<Candidate> CollectTopLevel_(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) const {
					std::vector<Candidate> all{};
					for (const auto& f : this->features_) {
						auto part = f->CollectChildren(grid, region);
						for (auto& c : part) {
							all.push_back(std::move(c));
						}
					}
					auto top = this->FilterTopLevel(all);
					LOG_DEBUG_D("top-level count: {}", static_cast<int>(top.size()));
					return top;
				}



				// 1) Оставить только top-level детей:
				// ребёнок А не top-level, если его bbox целиком лежит в любом subregion ребёнка B.
				std::vector<Candidate> FilterTopLevel(
					const std::vector<Candidate>& all
				) const {
					// Идея:
					// 1) Отсечь кандидатов, чьи bbox ПОЛНОСТЬЮ содержатся в bbox другого кандидата.
					//    Такие считаем вложенными — они не top-level.
					// 2) При равных bbox оставляем первый по порядку (стабильность).
					//
					// Сортировать не обязательно — достаточно одного O(n^2) прохода:
					// для каждого i проверяем, есть ли j != i, такой что bbox[j] содержит bbox[i].
					// Если да то i не топ.

					std::vector<Candidate> out{};
					out.reserve(all.size());

					for (size_t i = 0; i < all.size(); ++i) {
						const auto& ci = all[i];
						bool contained = false;

						for (size_t j = 0; j < all.size(); ++j) {
							if (i == j) {
								continue;
							}

							const auto& cj = all[j];

							// Если bbox[j] строго содержит bbox[i] то i не top-level
							if (cj.bbox.ContainsSubregion(ci.bbox)) {
								contained = true;
								break;
							}
						}

						if (!contained) {
							out.push_back(ci);
						}
					}

					// По желанию можно ещё убрать «дубликаты bbox», оставив первый
					// (обычно не требуется, но иногда полезно)
					{
						std::vector<Candidate> dedup{};
						dedup.reserve(out.size());

						auto same_bbox = [](const Model::Region& a, const Model::Region& b) {
							return a.Top() == b.Top()
								&& a.Bottom() == b.Bottom()
								&& a.Left() == b.Left()
								&& a.Right() == b.Right();
							};

						for (size_t i = 0; i < out.size(); ++i) {
							bool seen = false;
							for (size_t k = 0; k < i; ++k) {
								if (same_bbox(out[k].bbox, out[i].bbox)) {
									seen = true;
									break;
								}
							}
							if (!seen) {
								dedup.push_back(out[i]);
							}
						}

						out.swap(dedup);
					}

					return out;
				}


				// 2) Построить локальный skip = наследованный + полосы баров top-уровня
				std::unordered_map<int, std::vector<Model::SpanX>> BuildSkip_(
					const Model::Region& region,
					const std::vector<Candidate>& top,
					const std::unordered_map<int, std::vector<Model::SpanX>>& inherited
				) const {
					std::unordered_map<int, std::vector<Model::SpanX>> skip = inherited;
					for (const auto& ch : top) {
						ch.owner->AppendSkips(region, ch.id, skip); // добавит только bar
					}
					for (auto& kv : skip) {
						this->MergeRowSpans(kv.second);
					}
					return skip;
				}

				// Сливает перекрывающиеся или примыкающие (пересечения и касания) интервалы [x1..x2]
				// в пределах одной строки.
				void MergeRowSpans(std::vector<Model::SpanX>& spans) const {
					if (spans.empty()) {
						return;
					}
					std::sort(spans.begin(), spans.end(),
						[](const auto& a, const auto& b) {
							if (a.x1 != b.x1) {
								return a.x1 < b.x1;
							}
							return a.x2 < b.x2;
						});
					std::vector<Model::SpanX> merged{};
					merged.reserve(spans.size());

					Model::SpanX cur = spans.front();
					for (size_t i = 1; i < spans.size(); ++i) {
						if (spans[i].x1 <= (cur.x2 + 1)) {
							cur.x2 = std::max(cur.x2, spans[i].x2);
						}
						else {
							merged.push_back(cur);
							cur = spans[i];
						}
					}
					merged.push_back(cur);
					spans.swap(merged);
				}


				// 4) Занять регионы собранных фич на текущем уровне (чтобы не было «хвостов»)
				void OccupyFeatureRegions_(
					const Model::Region& region,
					const std::vector<std::unique_ptr<Model::INode>>& feats,
					std::unordered_map<int, std::vector<Model::SpanX>>& skip
				) const {
					for (const auto& node : feats) {
						const auto r = node->GetRegion();
						this->AddRegionToSkip(region, r, skip);
					}
					for (auto& kv : skip) {
						this->MergeRowSpans(kv.second);
					}
				}

				// Добавить регион r в skip (с клиппингом по region)
				void AddRegionToSkip(
					const Model::Region& region,
					const Model::Region& r,
					std::unordered_map<int, std::vector<Model::SpanX>>& skip
				) const {
					const int y1 = std::max(region.Top(), r.Top());
					const int y2 = std::min(region.Bottom(), r.Bottom());

					for (int y = y1; y <= y2; ++y) {
						const int left = std::max(region.Left(), r.Left());
						const int right = std::min(region.Right(), r.Right());
						if (left <= right) {
							skip[y].push_back(Model::SpanX{ left, right });
						}
					}
				}


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

						const auto allowedRanges = RegionWalker::AllowedRanges(
							region,
							rowSkips
						);

						{
							std::string line = std::format("  y={} allowed: ", y);
							for (const auto& a : allowedRanges) {
								line += std::format("[{}..{}] ", a.Left(), a.Right());
							}
							LOG_DEBUG_D("{}", line);
						}

						// Получаем «разрешённые» промежутки X в этой строке,
						// то есть области, не попадающие в skipRow.
						for (const auto& allowedRange : allowedRanges) {
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
									LOG_DEBUG_D(
										"    [y={}, x={}..{}] token='{}'",
										y,
										startX,
										x - 1,
										token
									);

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

					std::stable_sort(
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

			private:
				std::vector<std::unique_ptr<IRegionFeature>> features_;
			};
		}
	}
}