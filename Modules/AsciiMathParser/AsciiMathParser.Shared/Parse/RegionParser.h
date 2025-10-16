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
					auto children = this->CollectTopLevelChildBars(region, bars);
					this->SortBarsTopDown(children);

					// 2) рекурсивно построить Frac-узлы по детям
					std::vector<std::unique_ptr<Model::INode>> fracNodes{};

					for (const auto& b : children) {
						auto numNodes = this->ParseRegion(grid, b.numRegion, bars);
						auto denNodes = this->ParseRegion(grid, b.denRegion, bars);
						fracNodes.push_back(std::make_unique<Model::Frac>(
							Model::NodesGroup{ std::move(numNodes) },
							Model::NodesGroup{ std::move(denNodes) },
							b.barRegion.ToRegion()
						));
					}

					// 3) построить карту «skip» — зоны, которые нельзя токенизировать как символы
					const auto skipByRow = this->BuildSkipMap(region, bars);

					// 4) токенизация оставшегося текста в Symbol-узлы
					auto symNodes = this->TokenizeToSymbols(grid, region, skipByRow);

					// 5) слить Frac + Symbol по порядку чтения
					auto nodes = this->MergeByReadingOrder(
						std::move(fracNodes),
						std::move(symNodes)
					);

					return nodes;
				}

			private:
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
						if (b.barRegion.Left() < region.Left()) {
							continue;
						}
						if (b.barRegion.Right() > region.Right()) {
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
							return a.barRegion.Left() < b.barRegion.Left();
						}
					);
				}


				//  Формирует карту «запретных» зон (skipByRow):
				//  для каждой строки (y) в region список X-интервалов, которые нужно пропустить
				//  при токенизации. Туда попадают:
				//   • все линии черт (любых баров), пересекающих region по X,
				//   • окна Num/Den всех дочерних баров (чтобы не дублировать их контент).
				//  После сбора все интервалы на каждой строке сливаются в MergeRowSpans.
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
						const int L = std::max(region.Left(), b.barRegion.Left());
						const int R = std::min(region.Right(), b.barRegion.Right());
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
						for (int y = b.numRegion.Top(); y <= b.numRegion.Bottom(); ++y) {
							if (!region.ContainsY(y)) {
								continue;
							}
							const int L = std::max(region.Left(), b.numRegion.Left());
							const int R = std::min(region.Right(), b.numRegion.Right());
							if (L <= R) {
								skipByRow[y].push_back(Model::SpanX{ L, R });
							}
						}

						// DEN (клип по текущему region.cols)
						for (int y = b.denRegion.Top(); y <= b.denRegion.Bottom(); ++y) {
							if (!region.ContainsY(y)) {
								continue;
							}
							const int L = std::max(region.Left(), b.denRegion.Left());
							const int R = std::min(region.Right(), b.denRegion.Right());
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


				// Проходит по всем строкам region и формирует Symbol из подряд идущих
				// непробельных символов, пропуская все интервалы skipByRow.
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
						for (const auto& allowedRange : RegionParser::AllowedRanges(region, rowSkips)) {
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


				// Слияние по порядку чтения
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


				// Проверка: бар целиком «лежит» в регионе r
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
			};
		}
	}
}