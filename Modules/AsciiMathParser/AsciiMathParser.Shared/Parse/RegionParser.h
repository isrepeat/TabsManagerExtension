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
					// 1) Найти «детские» бары внутри региона и упорядочить
					auto children = this->CollectChildBars(
						region,
						bars
					);

					this->SortBarsTopDown(
						children
					);

					// 2) Сначала рекурсивно собрать Frac для каждого ребёнка
					std::vector<std::unique_ptr<Model::INode>> nodes{};

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

						Model::Num num{};
						for (auto& ptr : numNodes) {
							num.GetNodes().push_back(std::move(ptr));
						}

						Model::Den den{};
						for (auto& ptr : denNodes) {
							den.GetNodes().push_back(std::move(ptr));
						}

						nodes.push_back(
							std::make_unique<Model::Frac>(
								std::move(num),
								std::move(den)
							)
						);
					}

					// 3) Построить карту «skip» и дотокенизировать остатки как Symbol
					const auto skipByRow = this->BuildSkipMap(
						region,
						bars,
						children
					);

					auto tail = this->TokenizeRegionStrict(
						grid,
						region,
						skipByRow
					);

					for (auto& n : tail) {
						nodes.push_back(std::move(n));
					}

					return nodes;
				}

			private:
				// ---------- Шаг 1: выбор баров-детей ----------
				std::vector<Detect::FractionBar> CollectChildBars(
					const Model::Region& region,
					const std::vector<Detect::FractionBar>& allBars
				) const {
					std::vector<Detect::FractionBar> out{};
					out.reserve(allBars.size());

					for (const auto& b : allBars) {
						if (!region.ContainsY(b.y)) {
							continue;
						}

						if (b.x1 < region.cols.x1) {
							continue;
						}

						if (b.x2 > region.cols.x2) {
							continue;
						}

						out.push_back(b);
					}

					return out;
				}

				void SortBarsTopDown(
					std::vector<Detect::FractionBar>& bars
				) const {
					std::sort(
						bars.begin(),
						bars.end(),
						[](const Detect::FractionBar& a, const Detect::FractionBar& b) {
							if (a.y != b.y) {
								return a.y < b.y;
							}
							else {
								return a.x1 < b.x1;
							}
						}
					);
				}

				// ---------- Шаг 3: построение skip-карты ----------
				static void MergeRowSpans(
					std::vector<Model::SpanX>& spans
				) {
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

					for (const auto& s : spans) {
						if (merged.empty()) {
							merged.push_back(s);
						}
						else {
							auto& last = merged.back();

							if (s.x1 <= last.x2 + 1) {
								if (s.x2 > last.x2) {
									last.x2 = s.x2;
								}
							}
							else {
								merged.push_back(s);
							}
						}
					}

					spans = std::move(merged);
				}

				std::unordered_map<int, std::vector<Model::SpanX>> BuildSkipMap(
					const Model::Region& region,
					const std::vector<Detect::FractionBar>& allBars,
					const std::vector<Detect::FractionBar>& children
				) const {
					std::unordered_map<int, std::vector<Model::SpanX>> skipByRow{};

					// 3.1. Любая черта, чья строка попадает в region и по X пересекается с region, — в skip
					for (const auto& b : allBars) {
						if (!region.ContainsY(b.y)) {
							continue;
						}

						const int L = std::max(region.cols.x1, b.x1);
						const int R = std::min(region.cols.x2, b.x2);

						if (L <= R) {
							auto& vec = skipByRow[b.y];
							vec.push_back(Model::SpanX{ L, R });
						}
					}

					// 3.2. Для детей дополнительно исключаем их окна Num/Den полностью
					for (const auto& b : children) {
						// строка самой черты ребёнка
						{
							auto& vec = skipByRow[b.y];
							vec.push_back(Model::SpanX{ b.x1, b.x2 });
						}

						// строки числителя
						for (int y = b.numRegion.rows.y1; y <= b.numRegion.rows.y2; ++y) {
							auto& vec = skipByRow[y];
							vec.push_back(Model::SpanX{ b.numRegion.cols.x1, b.numRegion.cols.x2 });
						}

						// строки знаменателя
						for (int y = b.denRegion.rows.y1; y <= b.denRegion.rows.y2; ++y) {
							auto& vec = skipByRow[y];
							vec.push_back(Model::SpanX{ b.denRegion.cols.x1, b.denRegion.cols.x2 });
						}
					}

					// 3.3. Слить интервалы по каждой строке
					for (auto& kv : skipByRow) {
						RegionParser::MergeRowSpans(
							kv.second
						);
					}

					return skipByRow;
				}

				// ---------- Шаг 4: токенизация остатка ----------
				std::vector<std::unique_ptr<Model::INode>> TokenizeRegionStrict(
					const Model::AsciiGrid& grid,
					const Model::Region& region,
					const std::unordered_map<int, std::vector<Model::SpanX>>& skipByRow
				) const {
					std::vector<std::unique_ptr<Model::INode>> out{};

					for (int y = region.rows.y1; y <= region.rows.y2; ++y) {
						int x = region.cols.x1;

						const auto it = skipByRow.find(y);
						const bool hasSkips = (it != skipByRow.end());
						const std::vector<Model::SpanX>* spans = hasSkips ? &it->second : nullptr;

						while (x <= region.cols.x2) {
							RegionParser::JumpOverSkips(
								hasSkips,
								spans,
								x
							);

							while (x <= region.cols.x2 && grid.At(y, x) == ' ') {
								++x;

								RegionParser::JumpOverSkips(
									hasSkips,
									spans,
									x
								);
							}

							if (x > region.cols.x2) {
								break;
							}

							RegionParser::JumpOverSkips(
								hasSkips,
								spans,
								x
							);

							if (x > region.cols.x2) {
								break;
							}

							std::string token{};

							while (x <= region.cols.x2 && grid.At(y, x) != ' ') {
								if (RegionParser::IsInsideAnySpan(hasSkips, spans, x)) {
									break;
								}

								token.push_back(grid.At(y, x));
								++x;
							}

							if (!token.empty()) {
								out.push_back(
									std::make_unique<Model::Symbol>(
										std::move(token)
									)
								);
							}
						}
					}

					return out;
				}

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
			};
		}
	}
}