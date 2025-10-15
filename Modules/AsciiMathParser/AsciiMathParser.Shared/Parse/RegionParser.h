#pragma once
#include "../Detect/FractionBarDetector.h"
#include "../Model/Geometry.h"
#include "../Model/INode.h"
#include "../Model/Grid.h"
#include <unordered_set>
#include <unordered_map>
#include <algorithm>
#include <optional>
#include <memory>
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		namespace Parse {
			class RegionParser {
			public:
				std::vector<std::unique_ptr<Model::INode>> ParseRegion(
					const Model::AsciiGrid& grid,
					const Model::Region& region,
					const std::vector<Detect::FractionBar>& bars
				) const {
					// 1) собрать «детские» бары, лежащие внутри region, и отсортировать
					//auto children = this->CollectChildBars(region, bars);
					auto children = this->CollectTopLevelChildBars(region, bars);

					this->SortBarsTopDown(children);

					// 2) рекурсивно построить Frac-узлы по детям
					std::vector<NodeItem> fracItems{};

					for (const auto& b : children) {
						auto numNodes = this->ParseRegion(
							grid,
							b.numRegion,
							bars
						);

						auto denNodes = this->ParseRegion(
							grid,
							b.denRegion,
							bars
						);

						NodeItem it{
							.y = b.barRegion.y,
							.x = b.barRegion.cols.x1,
							.node = std::make_unique<Model::Frac>(
								Model::NodesGroup{ std::move(numNodes) },
								Model::NodesGroup{ std::move(denNodes) },
								b.barRegion.ToRegion()
							)
						};

						fracItems.push_back(std::move(it));
					}

					// 3) построить карту «skip» — зоны, которые нельзя токенизировать как символы
					const auto skipByRow = this->BuildSkipMap(
						region,
						bars
					);

					// 4) токенизация оставшегося текста в Symbol-узлы
					auto symItems = this->TokenizeToSymbolItems(
						grid,
						region,
						skipByRow
					);

					// 5) слить Frac + Symbol по порядку чтения
					auto nodes = this->MergeByReadingOrder(
						std::move(fracItems),
						std::move(symItems)
					);

					return nodes;
				}

			private:
				// ----- служебные записи для мерджа -----
				struct NodeItem {
					int y;
					int x;
					std::unique_ptr<Model::INode> node;
				};

				struct SymbolItem {
					int y;
					int xStart;
					std::string text;
				};


				// Возвращает список баров, которые физически находятся внутри region:
				// их строка черты (y) лежит в пределах region.rows,
				// а диапазон X полностью вписывается в region.cols.
				std::vector<Detect::FractionBar> CollectChildBars(
					const Model::Region& region,
					const std::vector<Detect::FractionBar>& allBars
				) const {
					std::vector<Detect::FractionBar> out{};
					out.reserve(allBars.size());

					for (const auto& b : allBars) {
						if (!region.ContainsY(b.barRegion.y)) {
							continue;
						}
						if (b.barRegion.cols.x1 < region.cols.x1) {
							continue;
						}
						if (b.barRegion.cols.x2 > region.cols.x2) {
							continue;
						}
						out.push_back(b);
					}

					return out;
				}


				// Собрать ВСЕ бары внутри region, затем отфильтровать только верхнеуровневые:
				// те, что не попадают внутрь num/den какого-либо другого из «внутренних».
				std::vector<Detect::FractionBar> CollectTopLevelChildBars(
					const Model::Region& region,
					const std::vector<Detect::FractionBar>& allBars
				) const {
					std::vector<Detect::FractionBar> inside{};
					inside.reserve(allBars.size());

					// 1) все бары целиком внутри region
					for (const auto& b : allBars) {
						if (RegionParser::IsBarInsideRegion(b, region)) {
							inside.push_back(b);
						}
					}

					// 2) оставить только те, которые не находятся в num/den других из inside
					std::vector<Detect::FractionBar> top{};
					top.reserve(inside.size());

					for (std::size_t i = 0; i < inside.size(); ++i) {
						const auto& bi = inside[i];
						bool nested = false;

						for (std::size_t j = 0; j < inside.size(); ++j) {
							if (i == j) {
								continue;
							}
							const auto& bj = inside[j];

							if (RegionParser::IsBarInsideRegion(bi, bj.numRegion) ||
								RegionParser::IsBarInsideRegion(bi, bj.denRegion)) {
								nested = true;
								break;
							}
						}

						if (!nested) {
							top.push_back(bi);
						}
					}

					return top;
				}


				//  Сортирует бары сверху вниз (по y) и слева направо (по x1),
				//  чтобы рекурсия шла в стабильном порядке.
				void SortBarsTopDown(std::vector<Detect::FractionBar>& bars) const {
					std::sort(
						bars.begin(),
						bars.end(),
						[](const Detect::FractionBar& a, const Detect::FractionBar& b) {
							if (a.barRegion.y != b.barRegion.y) {
								return a.barRegion.y < b.barRegion.y;
							}
							else {
								return a.barRegion.cols.x1 < b.barRegion.cols.x1;
							}
						}
					);
				}


				////  Формирует карту «запретных» зон (skipByRow):
				////  для каждой строки (y) в region список X-интервалов, которые нужно пропустить
				////  при токенизации. Туда попадают:
				////   • все линии черт (любых баров), пересекающих region по X,
				////   • окна Num/Den всех дочерних баров (чтобы не дублировать их контент).
				////  После сбора все интервалы на каждой строке сливаются в MergeRowSpans.
				//std::unordered_map<int, std::vector<Model::SpanX>> BuildSkipMap(
				//	const Model::Region& region,
				//	const std::vector<Detect::FractionBar>& allBars,
				//	const std::vector<Detect::FractionBar>& children
				//) const {
				//	std::unordered_map<int, std::vector<Model::SpanX>> skipByRow{};

				//	// 3.1 любые черты в пределах region по Y + пересечение по X
				//	for (const auto& b : allBars) {
				//		const int y = b.barRegion.y;
				//		if (!region.ContainsY(y)) {
				//			continue;
				//		}
				//		const int L = std::max(region.cols.x1, b.barRegion.cols.x1);
				//		const int R = std::min(region.cols.x2, b.barRegion.cols.x2);
				//		if (L <= R) {
				//			auto& vec = skipByRow[y];
				//			vec.push_back(Model::SpanX{ L, R });
				//		}
				//	}

				//	// 3.2 дочерние окна: бар, num, den — целиком в skip
				//	for (const auto& b : children) {
				//		// строка бара
				//		{
				//			auto& vec = skipByRow[b.barRegion.y];
				//			vec.push_back(Model::SpanX{ b.barRegion.cols.x1, b.barRegion.cols.x2 });
				//		}

				//		// строки числителя
				//		for (int y = b.numRegion.rows.y1; y <= b.numRegion.rows.y2; ++y) {
				//			auto& vec = skipByRow[y];
				//			vec.push_back(Model::SpanX{ b.numRegion.cols.x1, b.numRegion.cols.x2 });
				//		}

				//		// строки знаменателя
				//		for (int y = b.denRegion.rows.y1; y <= b.denRegion.rows.y2; ++y) {
				//			auto& vec = skipByRow[y];
				//			vec.push_back(Model::SpanX{ b.denRegion.cols.x1, b.denRegion.cols.x2 });
				//		}
				//	}

				//	// 3.3 слить интервалы на каждой строке
				//	for (auto& kv : skipByRow) {
				//		RegionParser::MergeRowSpans(
				//			kv.second
				//		);
				//	}

				//	return skipByRow;
				//}


				std::unordered_map<int, std::vector<Model::SpanX>> BuildSkipMap(
					const Model::Region& region,
					const std::vector<Detect::FractionBar>& allBars
				) const {
					std::unordered_map<int, std::vector<Model::SpanX>> skipByRow{};

					// 1) Все строки черт, пересекающие текущий region по X и попадающие по Y.
					for (const auto& b : allBars) {
						const int y = b.barRegion.y;
						if (!region.ContainsY(y)) {
							continue;
						}
						const int L = std::max(region.cols.x1, b.barRegion.cols.x1);
						const int R = std::min(region.cols.x2, b.barRegion.cols.x2);
						if (L <= R) {
							skipByRow[y].push_back(Model::SpanX{ L, R });
						}
					}

					// 2) Окна NUM/DEN ДЛЯ ВСЕХ баров, которые "лежат" внутри текущего region.
					//    Это гарантирует, что контент вложенных дробей не будет токенизироваться на этом уровне.
					for (const auto& b : allBars) {
						if (!RegionParser::IsBarInsideRegion(b, region)) {
							continue;
						}

						// NUM (клип по текущему region.cols)
						for (int y = b.numRegion.rows.y1; y <= b.numRegion.rows.y2; ++y) {
							if (!region.ContainsY(y)) {
								continue;
							}
							const int L = std::max(region.cols.x1, b.numRegion.cols.x1);
							const int R = std::min(region.cols.x2, b.numRegion.cols.x2);
							if (L <= R) {
								skipByRow[y].push_back(Model::SpanX{ L, R });
							}
						}

						// DEN (клип по текущему region.cols)
						for (int y = b.denRegion.rows.y1; y <= b.denRegion.rows.y2; ++y) {
							if (!region.ContainsY(y)) {
								continue;
							}
							const int L = std::max(region.cols.x1, b.denRegion.cols.x1);
							const int R = std::min(region.cols.x2, b.denRegion.cols.x2);
							if (L <= R) {
								skipByRow[y].push_back(Model::SpanX{ L, R });
							}
						}
					}

					// 3) Слить интервалы по строкам.
					for (auto& kv : skipByRow) {
						RegionParser::MergeRowSpans(kv.second);
					}

					return skipByRow;
				}


				//  Проходит по всем строкам region и формирует Symbol из подряд идущих
				//  непробельных символов, пропуская все интервалы skipByRow.
				//  Использует JumpOverSkips для прыжков через «стены».
				std::vector<SymbolItem> TokenizeToSymbolItems(
					const Model::AsciiGrid& grid,
					const Model::Region& region,
					const std::unordered_map<int, std::vector<Model::SpanX>>& skipByRow
				) const {
					std::vector<SymbolItem> items{};

					for (int y = region.rows.y1; y <= region.rows.y2; ++y) {
						int x = region.cols.x1;

						const auto it = skipByRow.find(y);
						const bool hasSpans = (it != skipByRow.end());
						const std::vector<Model::SpanX>* spans = hasSpans ? &it->second : nullptr;

						while (x <= region.cols.x2) {
							RegionParser::JumpOverSkips(
								hasSpans,
								spans,
								x
							);

							while (x <= region.cols.x2 && grid.At(x, y) == ' ') {
								++x;

								RegionParser::JumpOverSkips(
									hasSpans,
									spans,
									x
								);
							}

							if (x > region.cols.x2) {
								break;
							}

							RegionParser::JumpOverSkips(
								hasSpans,
								spans,
								x
							);

							if (x > region.cols.x2) {
								break;
							}

							const int startX = x;
							std::string token{};

							while (x <= region.cols.x2 && grid.At(x, y) != ' ') {
								if (RegionParser::IsInsideAnySpan(hasSpans, spans, x)) {
									break;
								}

								token.push_back(grid.At(x, y));
								++x;
							}

							if (!token.empty()) {
								items.push_back(SymbolItem{
									.y = y,
									.xStart = startX,
									.text = std::move(token)
									});
							}
						}
					}

					return items;
				}


				// ---------- шаг 5: слияние по порядку чтения ----------
				std::vector<std::unique_ptr<Model::INode>> MergeByReadingOrder(
					std::vector<NodeItem> fracItems,
					std::vector<SymbolItem> symItems
				) const {
					// преобразуем SymbolItem → NodeItem
					std::vector<NodeItem> all{};
					all.reserve(fracItems.size() + symItems.size());

					for (auto& it : fracItems) {
						all.push_back(NodeItem{
							.y = it.y,
							.x = it.x,
							.node = std::move(it.node)
							});
					}

					for (auto& s : symItems) {
						auto node = std::make_unique<Model::Symbol>(
							std::move(s.text),
							s.y,
							s.xStart
						);

						all.push_back(NodeItem{
							.y = s.y,
							.x = s.xStart,
							.node = std::move(node)
							});
					}

					std::sort(
						all.begin(),
						all.end(),
						[](const NodeItem& a, const NodeItem& b) {
							if (a.y != b.y) {
								return a.y < b.y;
							}
							else {
								return a.x < b.x;
							}
						}
					);

					std::vector<std::unique_ptr<Model::INode>> out{};
					out.reserve(all.size());

					for (auto& it : all) {
						out.push_back(std::move(it.node));
					}

					return out;
				}


				//  Сливает перекрывающиеся или примыкающие интервалы на одной строке:
				//  [3..8], [7..10] → [3..10]. Это делает карту skip компактной и
				//  уменьшает количество прыжков токенайзера.
				static void MergeRowSpans(std::vector<Model::SpanX>& spans) {
					std::sort(
						spans.begin(),
						spans.end(),
						[](const Model::SpanX& a, const Model::SpanX& b) {
							if (a.x1 != b.x1) {
								return a.x1 < b.x1;
							}
							else {
								return a.x2 < b.x2;
							}
						}
					);

					std::vector<Model::SpanX> merged{};
					std::optional<std::pair<int, int>> cur{};

					for (const auto& s : spans) {
						if (!cur.has_value()) {
							cur = std::make_pair(s.x1, s.x2);
						}
						else {
							auto& [cx1, cx2] = *cur;
							if (s.x1 <= cx2 + 1) {
								if (s.x2 > cx2) {
									cx2 = s.x2;
								}
							}
							else {
								merged.push_back(Model::SpanX{ cx1, cx2 });
								cur = std::make_pair(s.x1, s.x2);
							}
						}
					}

					if (cur.has_value()) {
						merged.push_back(Model::SpanX{ cur->first, cur->second });
					}

					spans = std::move(merged);
				}


				//  Если текущая позиция x находится внутри любого интервала skip,
				//  перепрыгивает на первую «безопасную» колонку справа (x = span.x2 + 1).
				static void JumpOverSkips(
					bool hasSpans,
					const std::vector<Model::SpanX>* spans,
					int& x
				) {
					if (!hasSpans) {
						return;
					}

					for (const auto& s : *spans) {
						if (x >= s.x1 && x <= s.x2) {
							x = s.x2 + 1;
							return;
						}
					}
				}


				//  Проверяет, лежит ли x внутри какого-либо интервала skip.
				//  Используется для обрыва токена, если в середине встретилась «стена».
				static bool IsInsideAnySpan(
					bool hasSpans,
					const std::vector<Model::SpanX>* spans,
					int x
				) {
					if (!hasSpans) {
						return false;
					}

					for (const auto& s : *spans) {
						if (x >= s.x1 && x <= s.x2) {
							return true;
						}
					}
					return false;
				}

				// Проверка: бар целиком «лежит» в регионе r
				static bool IsBarInsideRegion(
					const Detect::FractionBar& b,
					const Model::Region& r
				) {
					if (!r.ContainsY(b.barRegion.y)) {
						return false;
					}
					if (!r.ContainsX(b.barRegion.cols.x1)) {
						return false;
					}
					if (!r.ContainsX(b.barRegion.cols.x2)) {
						return false;
					}
					return true;
				}
			};
		}
	}
}