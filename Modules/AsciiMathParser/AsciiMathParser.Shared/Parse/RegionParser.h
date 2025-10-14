//#pragma once
//#include "../Detect/FractionBarDetector.h"
//#include "../Model/Geometry.h"
//#include "../Model/INode.h"
//#include "../Model/Grid.h"
//#include <unordered_set>
//#include <unordered_map>
//#include <algorithm>
//#include <optional>
//#include <memory>
//#include <vector>
//
//namespace AsciiMathParser {
//	namespace Core {
//		namespace Parse {
//			class RegionParser {
//			public:
//				std::vector<std::unique_ptr<Model::INode>> ParseRegion(
//					const Model::AsciiGrid& grid,
//					const Model::Region& region,
//					const std::vector<Detect::FractionBar>& bars
//				) const {
//					// 1) Найти «детские» бары внутри региона и упорядочить
//					auto children = this->CollectChildBars(
//						region,
//						bars
//					);
//
//					this->SortBarsTopDown(
//						children
//					);
//
//					// 2) Сначала рекурсивно собрать Frac для каждого ребёнка
//					std::vector<NodeItem> fracItems{};
//
//					for (const auto& b : children) {
//						auto numNodes = this->ParseRegion(
//							grid,
//							b.numRegion,
//							bars
//						);
//
//						auto denNodes = this->ParseRegion(
//							grid,
//							b.denRegion,
//							bars
//						);
//
//						NodeItem it{
//							.y = b.y,
//							.x = b.x1,
//							.node = std::make_unique<Model::Frac>(
//								Model::NodesGroup{ std::move(numNodes) },
//								Model::NodesGroup{ std::move(denNodes) }
//							)
//						};
//
//						fracItems.push_back(
//							std::move(it)
//						);
//					}
//
//					// 3) построить карту «skip» — интервалы, где токенайзер не должен читать текст
//					const auto skipByRow = this->BuildSkipMap(
//						region,
//						bars,
//						children
//					);
//
//					// 4) токенизировать остатки как Symbol-элементы с якорями
//					auto symItems = this->TokenizeToSymbolItems(
//						grid,
//						region,
//						skipByRow
//					);
//
//					// 5) слить два потока по (y, x) и вернуть итог
//					auto nodes = this->MergeByReadingOrder(
//						std::move(fracItems),
//						std::move(symItems)
//					);
//
//					return nodes;
//				}
//
//			private:
//				// ---------- данные для мерджа ----------
//				struct NodeItem {
//					int y;
//					int x;
//					std::unique_ptr<Model::INode> node;
//				};
//
//				struct SymbolItem {
//					int y;
//					int xStart;
//					std::string text;
//				};
//
//
//				//  Возвращает список баров, которые физически находятся внутри region:
//				//  их строка черты (y) лежит в пределах region.rows,
//				//  а диапазон X полностью вписывается в region.cols.
//				std::vector<Detect::FractionBar> CollectChildBars(
//					const Model::Region& region,
//					const std::vector<Detect::FractionBar>& allBars
//				) const {
//					std::vector<Detect::FractionBar> out{};
//					out.reserve(allBars.size());
//
//					for (const auto& b : allBars) {
//						if (!region.ContainsY(b.y)) {
//							continue;
//						}
//						if (b.x1 < region.cols.x1) {
//							continue;
//						}
//						if (b.x2 > region.cols.x2) {
//							continue;
//						}
//						out.push_back(b);
//					}
//
//					return out;
//				}
//
//
//				//  Сортирует бары сверху вниз (по y) и слева направо (по x1),
//				//  чтобы рекурсия шла в стабильном порядке.
//				void SortBarsTopDown(std::vector<Detect::FractionBar>& bars) const {
//					std::sort(
//						bars.begin(),
//						bars.end(),
//						[](const Detect::FractionBar& a, const Detect::FractionBar& b) {
//							if (a.y != b.y) {
//								return a.y < b.y;
//							}
//							else {
//								return a.x1 < b.x1;
//							}
//						}
//					);
//				}
//
//
//				//  Формирует карту «запретных» зон (skipByRow):
//				//  для каждой строки (y) в region список X-интервалов, которые нужно пропустить
//				//  при токенизации. Туда попадают:
//				//   • все линии черт (любых баров), пересекающих region по X,
//				//   • окна Num/Den всех дочерних баров (чтобы не дублировать их контент).
//				//  После сбора все интервалы на каждой строке сливаются в MergeRowSpans.
//				std::unordered_map<int, std::vector<Model::SpanX>> BuildSkipMap(
//					const Model::Region& region,
//					const std::vector<Detect::FractionBar>& allBars,
//					const std::vector<Detect::FractionBar>& children
//				) const {
//					std::unordered_map<int, std::vector<Model::SpanX>> skipByRow{};
//
//					// 3.1. Любая черта, чья строка попадает в region и по X пересекается с region, — в skip
//					for (const auto& b : allBars) {
//						if (!region.ContainsY(b.y)) {
//							continue;
//						}
//
//						const int L = std::max(region.cols.x1, b.x1);
//						const int R = std::min(region.cols.x2, b.x2);
//
//						if (L <= R) {
//							auto& vec = skipByRow[b.y];
//							vec.push_back(Model::SpanX{ L, R });
//						}
//					}
//
//					// 3.2. Для детей дополнительно исключаем их окна Num/Den полностью
//					for (const auto& b : children) {
//						// строка самой черты ребёнка
//						{
//							auto& vec = skipByRow[b.y];
//							vec.push_back(Model::SpanX{ b.x1, b.x2 });
//						}
//
//						// строки числителя
//						for (int y = b.numRegion.rows.y1; y <= b.numRegion.rows.y2; ++y) {
//							auto& vec = skipByRow[y];
//							vec.push_back(Model::SpanX{ b.numRegion.cols.x1, b.numRegion.cols.x2 });
//						}
//
//						// строки знаменателя
//						for (int y = b.denRegion.rows.y1; y <= b.denRegion.rows.y2; ++y) {
//							auto& vec = skipByRow[y];
//							vec.push_back(Model::SpanX{ b.denRegion.cols.x1, b.denRegion.cols.x2 });
//						}
//					}
//
//					// 3.3. Слить интервалы по каждой строке
//					for (auto& kv : skipByRow) {
//						RegionParser::MergeRowSpans(
//							kv.second
//						);
//					}
//
//					return skipByRow;
//				}
//
//
//				//  Проходит по всем строкам region и формирует Symbol из подряд идущих
//				//  непробельных символов, пропуская все интервалы skipByRow.
//				//  Использует JumpOverSkips для прыжков через «стены».
//				std::vector<SymbolItem> TokenizeToSymbolItems(
//					const Model::AsciiGrid& grid,
//					const Model::Region& region,
//					const std::unordered_map<int, std::vector<Model::SpanX>>& skipByRow
//				) const {
//					std::vector<SymbolItem> items{};
//
//					for (int y = region.rows.y1; y <= region.rows.y2; ++y) {
//						int x = region.cols.x1;
//
//						const auto it = skipByRow.find(y);
//						const bool hasSpans = (it != skipByRow.end());
//						const std::vector<Model::SpanX>* spans = hasSpans ? &it->second : nullptr;
//
//						while (x <= region.cols.x2) {
//							RegionParser::JumpOverSkips(
//								hasSpans,
//								spans,
//								x
//							);
//
//							while (x <= region.cols.x2 && grid.At(y, x) == ' ') {
//								++x;
//
//								RegionParser::JumpOverSkips(
//									hasSpans,
//									spans,
//									x
//								);
//							}
//
//							if (x > region.cols.x2) {
//								break;
//							}
//
//							RegionParser::JumpOverSkips(
//								hasSpans,
//								spans,
//								x
//							);
//
//							if (x > region.cols.x2) {
//								break;
//							}
//
//							const int startX = x;
//							std::string token{};
//
//							while (x <= region.cols.x2 && grid.At(y, x) != ' ') {
//								if (RegionParser::IsInsideAnySpan(hasSpans, spans, x)) {
//									break;
//								}
//
//								token.push_back(grid.At(y, x));
//								++x;
//							}
//
//							if (!token.empty()) {
//								items.push_back(SymbolItem{
//									.y = y,
//									.xStart = startX,
//									.text = std::move(token)
//									});
//							}
//						}
//					}
//
//					return items;
//				}
//
//
//				// ---------- шаг 5: слияние по порядку чтения ----------
//				std::vector<std::unique_ptr<Model::INode>> MergeByReadingOrder(
//					std::vector<NodeItem> fracItems,
//					std::vector<SymbolItem> symItems
//				) const {
//					// преобразуем SymbolItem → NodeItem
//					std::vector<NodeItem> all{};
//					all.reserve(fracItems.size() + symItems.size());
//
//					for (auto& it : fracItems) {
//						all.push_back(NodeItem{
//							.y = it.y,
//							.x = it.x,
//							.node = std::move(it.node)
//							});
//					}
//
//					for (auto& s : symItems) {
//						auto node = std::make_unique<Model::Symbol>(
//							std::move(s.text),
//							s.y,
//							s.xStart
//						);
//
//						all.push_back(NodeItem{
//							.y = s.y,
//							.x = s.xStart,
//							.node = std::move(node)
//							});
//					}
//
//					std::sort(
//						all.begin(),
//						all.end(),
//						[](const NodeItem& a, const NodeItem& b) {
//							if (a.y != b.y) {
//								return a.y < b.y;
//							}
//							else {
//								return a.x < b.x;
//							}
//						}
//					);
//
//					std::vector<std::unique_ptr<Model::INode>> out{};
//					out.reserve(all.size());
//
//					for (auto& it : all) {
//						out.push_back(std::move(it.node));
//					}
//
//					return out;
//				}
//
//
//				//  Сливает перекрывающиеся или примыкающие интервалы на одной строке:
//				//  [3..8], [7..10] → [3..10]. Это делает карту skip компактной и
//				//  уменьшает количество прыжков токенайзера.
//				static void MergeRowSpans(std::vector<Model::SpanX>& spans) {
//					std::sort(
//						spans.begin(),
//						spans.end(),
//						[](const Model::SpanX& a, const Model::SpanX& b) {
//							if (a.x1 != b.x1) {
//								return a.x1 < b.x1;
//							}
//							else {
//								return a.x2 < b.x2;
//							}
//						}
//					);
//
//					std::vector<Model::SpanX> merged{};
//
//					for (const auto& s : spans) {
//						if (merged.empty()) {
//							merged.push_back(s);
//						}
//						else {
//							auto& last = merged.back();
//
//							if (s.x1 <= last.x2 + 1) {
//								if (s.x2 > last.x2) {
//									last.x2 = s.x2;
//								}
//							}
//							else {
//								merged.push_back(s);
//							}
//						}
//					}
//
//					spans = std::move(merged);
//				}
//
//
//				//  Если текущая позиция x находится внутри любого интервала skip,
//				//  перепрыгивает на первую «безопасную» колонку справа (x = span.x2 + 1).
//				static void JumpOverSkips(
//					bool hasSpans,
//					const std::vector<Model::SpanX>* spans,
//					int& x
//				) {
//					if (!hasSpans) {
//						return;
//					}
//
//					for (const auto& s : *spans) {
//						if (x >= s.x1 && x <= s.x2) {
//							x = s.x2 + 1;
//							return;
//						}
//					}
//				}
//
//
//				//  Проверяет, лежит ли x внутри какого-либо интервала skip.
//				//  Используется для обрыва токена, если в середине встретилась «стена».
//				static bool IsInsideAnySpan(
//					bool hasSpans,
//					const std::vector<Model::SpanX>* spans,
//					int x
//				) {
//					if (!hasSpans) {
//						return false;
//					}
//
//					for (const auto& s : *spans) {
//						if (x >= s.x1 && x <= s.x2) {
//							return true;
//						}
//					}
//
//					return false;
//				}
//			};
//		}
//	}
//}