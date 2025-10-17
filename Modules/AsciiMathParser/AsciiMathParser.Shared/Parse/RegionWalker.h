#pragma once
#include <Helpers/Logger.h>

#include "../Model/Geometry.h"
#include "../Model/Symbol.h"
#include "../Model/Script.h"
#include "../Model/Group.h"
#include "IRegionFeature.h"

#include <unordered_map>
#include <algorithm>
#include <optional>

namespace AsciiMathParser {
	namespace Core {
		namespace Parse {
			// RegionWalker — одношаговый «сканер» области ASCII-сетки.
			//
			// Идеология:
			// 1) Сначала просим все IRegionFeature обнаружить своих кандидатов (Candidate) в заданном Region.
			//    Каждый Candidate самодостаточен: содержит bbox, список subregions для рекурсии,
			//    «структурные» пропуски (например, линия дробной черты), и лямбду assembleFn,
			//    которая собирает финальный узел из разобранных поддеревьев.
			// 2) Отбираем только top-level кандидатов (без тех, чей bbox целиком лежит внутри чужого bbox).
			// 3) Формируем локальную карту пропусков (mapRowToSkipRanges) как:
			//      inherited (от родителя)
			//      + structural (от всех top-level кандидатов на этом уровне).
			//    Эти пропуски нужны, чтобы при рекурсии и при токенизации не ловить «полосы»,
			//    которые являются частью оформления фич (напр. бар дроби).
			// 4) Рекурсивно разбираем subregions каждого кандидата, затем вызываем его assembleFn,
			//    получаем готовые «сложные» узлы (complexNodes).
			// 5) Занимаем их регионы в текущей карте пропусков (OccupyFeatureRegions),
			//    чтобы при токенизации на этом уровне не появились «хвосты» символов, перекрывающие уже собранные узлы.
			// 6) Всё, что не закрыто пропусками — токенизируем как «простые» символы.
			// 7) Склеиваем сложные и простые узлы и сортируем по порядку чтения (слева-направо, сверху-вниз).
			//
			class RegionWalker {
			public:
				explicit RegionWalker(
					std::vector<std::ex::unique_ptr<IRegionFeature>>&& features
				)
					: features{ std::move(features) } {
				}

				std::vector<std::ex::unique_ptr<Model::INode>> ParseRegion(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) const {
					std::unordered_map<int, std::vector<Model::SpanX>> mapRowToSkipRanges{};
					return this->ParseRegion(grid, region, mapRowToSkipRanges);
				}

				std::vector<std::ex::unique_ptr<Model::INode>> ParseRegion(
					const Model::AsciiGrid& grid,
					const Model::Region& region,
					const std::unordered_map<int, std::vector<Model::SpanX>>& inheritedMapRowToSkipRanges
				) const {
                    LOG_FUNCTION_SCOPE("ParseRegion(): bbox=[x:{}..{} y:{}..{}]",
                        region.Left(),
                        region.Right(),
                        region.Top(),
                        region.Bottom()
                    );

					const auto topCandidates = this->CollectTopLevel(grid, region);
					auto mapRowToSkipRanges = this->BuildRowSkipRangesMap(
						region,
						topCandidates,
						inheritedMapRowToSkipRanges
					);

					std::vector<std::ex::unique_ptr<Model::INode>> complexNodes{};
					complexNodes.reserve(topCandidates.size());

					for (const auto& topCandidate : topCandidates) {
						std::vector<std::vector<std::ex::unique_ptr<Model::INode>>> subtrees{};
						subtrees.reserve(topCandidate.subregions.size());

						for (const auto& sub : topCandidate.subregions) {
							auto subNodes = this->ParseRegion(
								grid,
								sub,
								mapRowToSkipRanges
							);
							subtrees.push_back(std::move(subNodes));
						}
						complexNodes.push_back(topCandidate.assembleFn(std::move(subtrees)));
					}

					this->OccupyFeatureRegions(
						region,
						complexNodes,
						mapRowToSkipRanges
					);

					auto symbolsNodes = this->TokenizeToSymbols(
						grid,
						region,
						mapRowToSkipRanges
					);

					auto nodes = this->MergeByReadingOrder(
						std::move(complexNodes),
						std::move(symbolsNodes)
					);

                    nodes = this->AssembleInlineScripts(
                        std::move(nodes)
                    );

					return nodes;
				}

			private:
				// 1)
				std::vector<Candidate> CollectTopLevel(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) const {
					std::vector<Candidate> all;

					for (const auto& f : this->features) {
						auto part = f->CollectChildren(grid, region);
						for (auto& c : part) {
							all.push_back(std::move(c));
						}
					}
					return this->FilterTopLevel(all);
				}


				// Оставить только top-level детей:
				// ребёнок А не top-level, если его bbox целиком лежит в любом subregion ребёнка B.
				std::vector<Candidate> FilterTopLevel(
					const std::vector<Candidate>& all
				) const {
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

							// Если bbox[j] строго содержит bbox[i] то 'i' не top-level
							if (cj.bbox.ContainsSubregion(ci.bbox)) {
								contained = true;
								break;
							}
						}

						if (!contained) {
							out.push_back(ci);
						}
					}

					return out;
				}


				// 2) Собирает рабочую карту пропусков строк → X-интервалы.
				std::unordered_map<int, std::vector<Model::SpanX>> BuildRowSkipRangesMap(
					const Model::Region& region,
					const std::vector<Candidate>& topCandidates,
					const std::unordered_map<int, std::vector<Model::SpanX>>& inheritedMapRowToSkipRanges
				) const {					
					auto mapRowToSkipRanges = inheritedMapRowToSkipRanges;
					// mapRowToSkipRanges — это рабочая агрегированная карта пропусков текущего уровня. 
					// Она получается так:
					// - берём inheritedMapRowToSkipRanges от предка (что уже нельзя трогать наверху),
					// - добавляем все mapRowToSkipRangesStructural для текущих top - кандидатов.

					for (const auto& topCandidate : topCandidates) {
						for (const auto& [y, spans] : topCandidate.mapRowToSkipRangesStructural) {
							if (!region.ContainsY(y)) {
								continue;
							}
							
							for (const auto& s : spans) {
								// клип по текущему region
								const int L = std::max(region.Left(), s.x1);
								const int R = std::min(region.Right(), s.x2);
								if (L <= R) {
									mapRowToSkipRanges[y].push_back(Model::SpanX{ L, R });
								}
							}
						}
					}

					for (auto& kv : mapRowToSkipRanges) {
						this->MergeRowSpans(kv.second);
					}
					return mapRowToSkipRanges;
				}


				// Сливает перекрывающиеся или примыкающие интервалы [x1..x2] в пределах одной строки.
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


				// 3) Занять регионы собранных фич на текущем уровне (чтобы не было «хвостов»)
				void OccupyFeatureRegions(
					const Model::Region& region,
					const std::vector<std::ex::unique_ptr<Model::INode>>& feats,
					std::unordered_map<int, std::vector<Model::SpanX>>& mapRowToSkipRanges
				) const {
					for (const auto& node : feats) {
						const auto r = node->GetRegion();
						this->AddRegionToSkip(region, r, mapRowToSkipRanges);
					}
					for (auto& kv : mapRowToSkipRanges) {
						this->MergeRowSpans(kv.second);
					}
				}


				// Добавить регион r в skip (с клиппингом по region)
				void AddRegionToSkip(
					const Model::Region& region,
					const Model::Region& r,
					std::unordered_map<int, std::vector<Model::SpanX>>& mapRowToSkipRanges
				) const {
					const int y1 = std::max(region.Top(), r.Top());
					const int y2 = std::min(region.Bottom(), r.Bottom());

					for (int y = y1; y <= y2; ++y) {
						const int left = std::max(region.Left(), r.Left());
						const int right = std::min(region.Right(), r.Right());
						if (left <= right) {
							mapRowToSkipRanges[y].push_back(Model::SpanX{ left, right });
						}
					}
				}


				// 4) Всё, что не закрыто пропусками — токенизируем как «простые» символы.
				std::vector<std::ex::unique_ptr<Model::INode>> TokenizeToSymbols(
					const Model::AsciiGrid& grid,
					const Model::Region& region,
					const std::unordered_map<int, std::vector<Model::SpanX>>& mapRowToSkipRanges
				) const {
					std::vector<std::ex::unique_ptr<Model::INode>> out{};

					// Проходим все строки текущего региона сверху вниз.
					for (int y = region.Top(); y <= region.Bottom(); ++y) {

						// Получаем список X-интервалов, которые нужно пропустить на этой строке.
						// Если строка не имеет пропусков — используем пустой список.
						const auto it = mapRowToSkipRanges.find(y);
						static const std::vector<Model::SpanX> kEmpty{};
						const auto& skipRowRanges = (it == mapRowToSkipRanges.end() ? kEmpty : it->second);

						const auto rowAllowedRanges = this->GetRowAllowedRanges(
							region,
							skipRowRanges
						);

						{
							std::string line = std::format("  y={} allowed: ", y);
							for (const auto& a : rowAllowedRanges) {
								line += std::format("[{}..{}] ", a.Left(), a.Right());
							}
							LOG_DEBUG_D("{}", line);
						}

						// Получаем «разрешённые» промежутки X в этой строке,
						// то есть области, не попадающие в skipRow.
						for (const auto& allowedSpanX : rowAllowedRanges) {
							// Начинаем обход разрешённого диапазона слева направо.
							int x = allowedSpanX.Left();

                            // Идём по всем символам до конца разрешённого диапазона.
                            while (x <= allowedSpanX.Right()) {
                                // Пропускаем все пробелы подряд.
                                while (x <= allowedSpanX.Right() && grid.At(x, y) == ' ') {
                                    ++x;
                                }

                                if (x > allowedSpanX.Right()) {
                                    break;
                                }

                                // Если текущий символ — '^' или '_' — это всегда отдельный односимвольный токен.
                                const char ch = grid.At(x, y);
                                if (ch == '^' || ch == '_') {
                                    LOG_DEBUG_D("    [y={}, x={}] token='{}'", y, x, std::string(1, ch));

                                    std::string one{ ch };
                                    out.push_back(std::ex::make_unique_ex<Model::Symbol>(
                                        std::move(one),
                                        x,
                                        y
                                    ));

                                    ++x;
                                    continue;
                                }

                                // Иначе накапливаем «слово» до ближайшего пробела ИЛИ до '^'/'_'.
                                const int startX = x;
                                std::string token{};

                                while (x <= allowedSpanX.Right()) {
                                    const char c = grid.At(x, y);
                                    if (c == ' ' || c == '^' || c == '_') {
                                        break;
                                    }
                                    token.push_back(c);
                                    ++x;
                                }

                                if (!token.empty()) {
                                    LOG_DEBUG_D(
                                        "    [y={}, x={}..{}] token='{}'",
                                        y,
                                        startX,
                                        x - 1,
                                        token
                                    );

                                    out.push_back(std::ex::make_unique_ex<Model::Symbol>(
                                        std::move(token),
                                        startX,
                                        y
                                    ));
                                }

                                // Цикл не «съедает» '^'/'_' здесь — на следующей итерации они будут разобраны веткой выше
                            }
						}
					}

					return out;
				}


				// Разность по X: всё, что можно читать в строке 'y' внутри region (без skip-интервалов).
                std::vector<Model::SpanX> GetRowAllowedRanges(
                    const Model::Region& region,
                    const std::vector<Model::SpanX>& skipRowRanges
                ) const {
                    std::vector<Model::SpanX> allowed{};

                    // 1) Копия skip'ов, ОГРАНИЧЕННАЯ текущим region и отфильтрованная по пересечению
                    std::vector<Model::SpanX> clipped{};
                    clipped.reserve(skipRowRanges.size());

                    for (const auto& s : skipRowRanges) {
                        // нет пересечения с region -> пропускаем
                        if (s.Right() < region.Left() || s.Left() > region.Right()) {
                            continue;
                        }
                        // клип по границам region
                        Model::SpanX t{
                            std::max(s.Left(), region.Left()),
                            std::min(s.Right(), region.Right())
                        };
                        clipped.push_back(t);
                    }

                    if (clipped.empty()) {
                        allowed.push_back(Model::SpanX{ region.Left(), region.Right() });
                        return allowed;
                    }

                    // 2) Слить пересекающиеся/смежные пропуски в пределах region
                    std::sort(
                        clipped.begin(),
                        clipped.end(),
                        [](const Model::SpanX& a, const Model::SpanX& b) {
                            if (a.x1 != b.x1) {
                                return a.x1 < b.x1;
                            }
                            return a.x2 < b.x2;
                        }
                    );

                    std::vector<Model::SpanX> merged{};
                    merged.reserve(clipped.size());

                    Model::SpanX cur = clipped.front();
                    for (std::size_t i = 1; i < clipped.size(); ++i) {
                        if (clipped[i].x1 <= (cur.x2 + 1)) {
                            cur.x2 = std::max(cur.x2, clipped[i].x2);
                        }
                        else {
                            merged.push_back(cur);
                            cur = clipped[i];
                        }
                    }
                    merged.push_back(cur);

                    // 3) Построить allowed как дополнение merged внутри region
                    int x = region.Left();
                    for (const auto& skip : merged) {
                        if (skip.x1 > x) {
                            allowed.push_back(Model::SpanX{ x, skip.x1 - 1 });
                        }
                        x = skip.x2 + 1;
                        if (x > region.Right()) {
                            break;
                        }
                    }
                    if (x <= region.Right()) {
                        allowed.push_back(Model::SpanX{ x, region.Right() });
                    }

                    return allowed;
                }



				// 5) Объединяет сложные и простые узлы и сортирует их «по чтению»:
				// первично по Left(), вторично по Top().
				// Используем stable_sort, чтобы сохранять относительный порядок элементов
				std::vector<std::ex::unique_ptr<Model::INode>> MergeByReadingOrder(
					std::vector<std::ex::unique_ptr<Model::INode>> complexNodes,
					std::vector<std::ex::unique_ptr<Model::INode>> symbolsNodes
				) const {
					std::vector<std::ex::unique_ptr<Model::INode>> all{};
					all.reserve(complexNodes.size() + symbolsNodes.size());

					for (auto& p : complexNodes) {
						all.push_back(std::move(p));
					}
					for (auto& p : symbolsNodes) {
						all.push_back(std::move(p));
					}

					std::stable_sort(
						all.begin(),
						all.end(),
						[](const std::ex::unique_ptr<Model::INode>& a, const std::ex::unique_ptr<Model::INode>& b) {
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


                // Склеивает [Base] ^ {Group} [_ {Group}]
                // Условие: '^' должен идти ДО '_' (если оба есть).
                std::vector<std::ex::unique_ptr<Model::INode>> AssembleInlineScripts(
                    std::vector<std::ex::unique_ptr<Model::INode>> nodes
                ) const {
                    if (nodes.empty()) {
                        return nodes;
                    }

                    std::vector<std::ex::unique_ptr<Model::INode>> out{};
                    out.reserve(nodes.size());

                    for (std::size_t i = 0; i < nodes.size(); /* i изменяем внутри */) {
                        // База — любой узел на текущей позиции.
                        auto& nodeBase = nodes[i];

                        // Попытаемся распознать суффиксы супер/саб сразу справа от базы в той же строке.
                        std::optional<Model::NodesGroup> nodesGroupSup{};
                        std::optional<Model::NodesGroup> nodesGroupSub{};

                        std::size_t j = i + 1;

                        // Вспомогалка: если на позиции idx стоит Group — забрать ЕЁ ВНУТРЕННОСТЬ переносом и удалить Group.
                        auto tryTakeGroupFn = [&](std::size_t& idx) -> std::optional<Model::NodesGroup> {
                            if (idx >= nodes.size()) {
                                return std::nullopt;
                            }

                            //auto* group = dynamic_cast<Model::Group*>(nodes[idx].get());
                            auto group = nodes[idx].AsView<Model::Group>();
                            if (group == nullptr) {
                                return std::nullopt;
                            }

                            Model::NodesGroup nodesGroup = group->ReleaseInner();
                            nodes.erase(nodes.begin() + static_cast<std::ptrdiff_t>(idx));
                            return nodesGroup;
                            };

                        auto isSymbolWithCharFn = [](const std::ex::unique_ptr<Model::INode>& n, char ch) {
                            if (auto symbol = n.AsView<Model::Symbol>()) {
                                const auto& content = symbol->Content();
                                return content.size() == 1 && content[0] == ch;
                            }
                            return false;
                            };

                        // ===== СУПЕРСКРИПТ =====
                        // Правило: удаляем '^' ТОЛЬКО если сразу за ним идёт {Group} в той же строке.
                        if (j < nodes.size()) {
                            const bool sameRow = (nodes[j]->GetRegion().Top() == nodeBase->GetRegion().Top());
                            const bool isCaret = isSymbolWithCharFn(nodes[j], '^');

                            if (sameRow && isCaret) {
                                // Проверим, что следом действительно стоит Group; иначе — НЕ трогаем '^'.
                                if ((j + 1) < nodes.size() &&
                                    nodes[j + 1].Is<Model::Group>()
                                    ) {
                                    LOG_DEBUG_D("Assemble: found '^' with Group");

                                    // Удаляем '^'; теперь на позиции j стоит Group.
                                    nodes.erase(nodes.begin() + static_cast<std::ptrdiff_t>(j));

                                    if (auto grabbed = tryTakeGroupFn(j); grabbed.has_value()) {
                                        nodesGroupSup = std::move(grabbed);
                                    }
                                }
                                else {
                                    LOG_DEBUG_D("Assemble: '^' has no Group next — skip");
                                }
                            }
                        }

                        // ===== САБСКРИПТ =====
                        // Аналогично: '_' удаляем только если за ним немедленно идёт {Group} в той же строке.
                        if (j < nodes.size()) {
                            const bool sameRow = (nodes[j]->GetRegion().Top() == nodeBase->GetRegion().Top());
                            const bool isUnderscore = isSymbolWithCharFn(nodes[j], '_');

                            if (sameRow && isUnderscore) {
                                if ((j + 1) < nodes.size() &&
                                    nodes[j + 1].Is<Model::Group>()
                                    ) {
                                    LOG_DEBUG_D("Assemble: found '_' with Group");

                                    nodes.erase(nodes.begin() + static_cast<std::ptrdiff_t>(j));

                                    if (auto grabbed = tryTakeGroupFn(j); grabbed.has_value()) {
                                        nodesGroupSub = std::move(grabbed);
                                    }
                                }
                                else {
                                    LOG_DEBUG_D("Assemble: '_' has no Group next — skip");
                                }
                            }
                        }

                        // Если что-то из sup/sub нашлось — строим Script и «съедаем» base.
                        if (nodesGroupSup.has_value() ||
                            nodesGroupSub.has_value()
                            ) {
                            auto ownedBase = std::move(nodeBase);
                            nodes.erase(nodes.begin() + static_cast<std::ptrdiff_t>(i));

                            auto scriptNode = std::ex::make_unique_ex<Model::Script>(
                                std::move(ownedBase),
                                std::move(nodesGroupSup),
                                std::move(nodesGroupSub)
                            );

                            out.push_back(std::move(scriptNode));
                            // i не увеличиваем: на этой позиции уже стоит следующий элемент,
                            // цикл продолжит с текущего i.
                        }
                        else {
                            // Ничего не склеилось — переносим base как есть.
                            out.push_back(std::move(nodeBase));
                            nodes.erase(nodes.begin() + static_cast<std::ptrdiff_t>(i));
                            // i не увеличиваем по той же причине — на i сместился следующий узел.
                        }
                    }

                    // Перенесём хвост (на всякий случай) — обычно пусто.
                    for (auto& node : nodes) {
                        out.push_back(std::move(node));
                    }

                    // Стабилизируем порядок чтения (слева-направо, сверху-вниз).
                    std::stable_sort(
                        out.begin(),
                        out.end(),
                        [](const std::ex::unique_ptr<Model::INode>& nodeA,
                            const std::ex::unique_ptr<Model::INode>& nodeB) {
                                const auto rA = nodeA->GetRegion();
                                const auto rB = nodeB->GetRegion();

                                if (rA.Left() != rB.Left()) {
                                    return (rA.Left() < rB.Left());
                                }
                                else {
                                    return (rA.Top() < rB.Top());
                                }
                        }
                    );

                    return out;
                }

			private:
				std::vector<std::ex::unique_ptr<IRegionFeature>> features;
			};
		}
	}
}