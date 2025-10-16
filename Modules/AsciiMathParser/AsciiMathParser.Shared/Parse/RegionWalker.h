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

				std::vector<std::unique_ptr<Model::INode>> ParseRegion(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) const {
					LOG_FUNCTION_SCOPE(
						"ParseRegion(): bbox=[x:{}..{} y:{}..{}]",
						region.Left(),
						region.Right(),
						region.Top(),
						region.Bottom()
					);

					// 1) Сбор кандидатов от всех фич
					std::vector<PlannedChild> all{};
					for (const auto& f : this->features_) {
						auto v = f->CollectChildren(grid, region);

						LOG_DEBUG_D(
							"feature {} -> {} candidates",
							(const void*)f.get(),
							(int)v.size()
						);

						for (const auto& c : v) {
							LOG_DEBUG_D(
								"  cand id={} bbox=[{}..{} x {}..{}]; subregions={}",
								(unsigned long long)c.id,
								c.bbox.Left(),
								c.bbox.Right(),
								c.bbox.Top(),
								c.bbox.Bottom(),
								(int)c.subregions.size()
							);
						}

						for (auto& x : v) {
							all.push_back(std::move(x));
						}
					}

					// 2) Оставить только top-level
					auto top = this->FilterTopLevel(all);

					LOG_DEBUG_D(
						"top-level count: {}",
						(int)top.size()
					);

					for (const auto& c : top) {
						LOG_DEBUG_D(
							"  top id={} bbox=[{}..{} x {}..{}]",
							(unsigned long long)c.id,
							c.bbox.Left(),
							c.bbox.Right(),
							c.bbox.Top(),
							c.bbox.Bottom()
						);
					}

					// 3) Сформировать skip-карту под top-детей
					std::unordered_map<int, std::vector<Model::SpanX>> skip{};
					for (const auto& ch : top) {
						ch.owner->AppendSkips(region, ch.id, skip);
					}
					for (auto& kv : skip) {
						this->MergeRowSpans(kv.second);
					}

					LOG_DEBUG_D("skip map (merged):");
					for (const auto& kv : skip) {
						std::string line{};
						line += std::format("  y={}: ", kv.first);
						for (const auto& s : kv.second) {
							line += std::format("[{}..{}] ", s.Left(), s.Right());
						}
						LOG_DEBUG_D("{}", line);
					}

					// 4) Рекурсивно собрать поддеревья и «собрать» узлы у фичи
					std::vector<std::unique_ptr<Model::INode>> featureNodes{};
					featureNodes.reserve(top.size());

					for (const auto& ch : top) {
						LOG_DEBUG_D(
							"assemble child id={} subregions={}",
							(unsigned long long)ch.id,
							(int)ch.subregions.size()
						);

						std::vector<std::vector<std::unique_ptr<Model::INode>>> subtrees{};
						subtrees.reserve(ch.subregions.size());

						for (const auto& sub : ch.subregions) {
							LOG_DEBUG_D(
								"  recurse sub=[{}..{} x {}..{}]",
								sub.Left(),
								sub.Right(),
								sub.Top(),
								sub.Bottom()
							);

							auto subNodes = this->ParseRegion(
								grid,
								sub
							);

							LOG_DEBUG_D(
								"  sub parsed: {} nodes",
								(int)subNodes.size()
							);

							subtrees.push_back(std::move(subNodes));
						}
						featureNodes.push_back(
							ch.owner->Assemble(ch.id, std::move(subtrees))
						);
					}

					// 5) Оставшееся → символы
					auto symbolNodes = this->TokenizeToSymbols(grid, region, skip);

					// 6) Слить по порядку чтения
					auto merged = this->MergeByReadingOrder(
						std::move(featureNodes),
						std::move(symbolNodes)
					);

					LOG_DEBUG_D(
						"merged nodes: {}",
						(int)merged.size()
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


				// 1) Оставить только top-level детей:
				  // ребёнок А не top-level, если его bbox целиком лежит в любом subregion ребёнка B.
				std::vector<PlannedChild> FilterTopLevel(
					const std::vector<PlannedChild>& in
				) const {
					std::vector<PlannedChild> out{};
					out.reserve(in.size());

					for (std::size_t i = 0; i < in.size(); ++i) {
						bool nested = false;

						for (std::size_t j = 0; j < in.size(); ++j) {
							if (i == j) {
								continue;
							}

							for (const auto& owned : in[j].subregions) {
								if (owned.ContainsSubregion(in[i].bbox)) {
									nested = true;
									break;
								}
							}

							if (nested == true) {
								break;
							}
						}

						if (nested == false) {
							out.push_back(in[i]);
						}
					}

					// стабильный reading-order: сначала по X (Left), затем по Y (Top)
					std::sort(
						out.begin(),
						out.end(),
						[](const PlannedChild& a, const PlannedChild& b) {
							if (a.bbox.Left() != b.bbox.Left()) {
								return a.bbox.Left() < b.bbox.Left();
							}
							return a.bbox.Top() < b.bbox.Top();
						}
					);

					LOG_DEBUG_D("FilterTopLevel: in={}, out={}", (int)in.size(), (int)out.size());
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


				// 2) Слить/нормализовать X-интервалы (пересечения и касания).
				// Сливает перекрывающиеся или примыкающие интервалы на одной строке:
				// [3..8], [7..10] → [3..10]. Это делает карту skip компактной и
				// уменьшает количество прыжков токенайзера.
				static void MergeRowSpans(
					std::vector<Model::SpanX>& spans
				) {
					if (spans.empty() == true) {
						return;
					}

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
					merged.reserve(spans.size());

					Model::SpanX acc = spans.front();
					for (std::size_t i = 1; i < spans.size(); ++i) {
						const auto& s = spans[i];
						if (s.Left() <= acc.Right() + 1) {
							// пересекаются или соприкасаются → расширяем
							acc.x2 = std::max(acc.Right(), s.Right());
						}
						else {
							merged.push_back(acc);
							acc = s;
						}
					}
					merged.push_back(acc);

					spans = std::move(merged);
				}

			private:
				std::vector<std::unique_ptr<IRegionFeature>> features_;
			};
		}
	}
}