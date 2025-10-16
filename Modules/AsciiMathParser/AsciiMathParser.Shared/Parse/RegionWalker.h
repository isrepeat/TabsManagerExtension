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
					std::vector<std::unique_ptr<IRegionFeature>>&& features
				)
					: features{ std::move(features) } {
				}

				std::vector<std::unique_ptr<Model::INode>> ParseRegion(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) const {
					std::unordered_map<int, std::vector<Model::SpanX>> mapRowToSkipRanges{};
					return this->ParseRegion(grid, region, mapRowToSkipRanges);
				}

				std::vector<std::unique_ptr<Model::INode>> ParseRegion(
					const Model::AsciiGrid& grid,
					const Model::Region& region,
					const std::unordered_map<int, std::vector<Model::SpanX>>& inheritedMapRowToSkipRanges
				) const {
					const auto topCandidates = this->CollectTopLevel(grid, region);
					auto mapRowToSkipRanges = this->BuildRowSkipRangesMap(
						region,
						topCandidates,
						inheritedMapRowToSkipRanges
					);

					std::vector<std::unique_ptr<Model::INode>> complexNodes{};
					complexNodes.reserve(topCandidates.size());

					for (const auto& topCandidate : topCandidates) {
						std::vector<std::vector<std::unique_ptr<Model::INode>>> subtrees{};
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
					const std::vector<std::unique_ptr<Model::INode>>& feats,
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
				std::vector<std::unique_ptr<Model::INode>> TokenizeToSymbols(
					const Model::AsciiGrid& grid,
					const Model::Region& region,
					const std::unordered_map<int, std::vector<Model::SpanX>>& mapRowToSkipRanges
				) const {
					std::vector<std::unique_ptr<Model::INode>> out{};

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

								// Если после пропусков дошли до конца диапазона — строка исчерпана.
								if (x > allowedSpanX.Right()) {
									break;
								}

								// Здесь x указывает на первый непробельный символ токена.
								const int startX = x;
								std::string token{};

								// Собираем последовательность непробельных символов до конца токена
								// (пока не встретим пробел или конец диапазона).
								while (x <= allowedSpanX.Right() && grid.At(x, y) != ' ') {
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


				// Разность по X: всё, что можно читать в строке 'y' внутри region (без skip-интервалов).
				std::vector<Model::SpanX> GetRowAllowedRanges(
					const Model::Region& region,
					const std::vector<Model::SpanX>& skipRowRanges
				) const {
					std::vector<Model::SpanX> rowAllowedRanges{};

					if (skipRowRanges.empty()) {
						rowAllowedRanges.push_back(Model::SpanX{ region.Left(), region.Right() });
						return rowAllowedRanges;
					}

					int cur = region.Left();
					for (const auto& spanX : skipRowRanges) {
						if (spanX.Left() > cur) {
							rowAllowedRanges.push_back(Model::SpanX{ cur, spanX.Left() - 1 });
						}
						cur = std::max(cur, spanX.Right() + 1);
					}

					if (cur <= region.Right()) {
						rowAllowedRanges.push_back(Model::SpanX{ cur, region.Right() });
					}

					return rowAllowedRanges;
				}


				// 5) Объединяет сложные и простые узлы и сортирует их «по чтению»:
				// первично по Left(), вторично по Top().
				// Используем stable_sort, чтобы сохранять относительный порядок элементов
				std::vector<std::unique_ptr<Model::INode>> MergeByReadingOrder(
					std::vector<std::unique_ptr<Model::INode>> complexNodes,
					std::vector<std::unique_ptr<Model::INode>> symbolsNodes
				) const {
					std::vector<std::unique_ptr<Model::INode>> all{};
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

			private:
				std::vector<std::unique_ptr<IRegionFeature>> features;
			};
		}
	}
}