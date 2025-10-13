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
					// --- 1) Выбрать бары, чьи черты попадают внутрь region и полностью вписываются по X.
					std::vector<Detect::FractionBar> children{};
					children.reserve(bars.size());

					for (const auto& b : bars) {
						const bool yInside = region.ContainsY(b.y);

						if (!yInside) {
							continue;
						}

						if (b.x1 < region.cols.x1) {
							continue;
						}

						if (b.x2 > region.cols.x2) {
							continue;
						}

						children.push_back(b);
					}

					std::sort(
						children.begin(),
						children.end(),
						[](const Detect::FractionBar& a, const Detect::FractionBar& b) {
							if (a.y != b.y) {
								return a.y < b.y;
							}
							else {
								return a.x1 < b.x1;
							}
						}
					);

					// --- 2) Рекурсивно собрать Frac для каждого дочернего бара.
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

					// --- 3) Построим «skip»-интервалы по строкам -------------------------------
					// Смысл: токенайзер не должен читать символы на строках черт, которые
					// попадают в текущий region, даже если такая черта не является child.
					// Плюс исключаем окна Num/Den дочерних дробей.

					std::unordered_map<int, std::vector<Model::SpanX>> skipByRow{};

					// 3.1. Любые черты, чья линия y лежит в пределах region.rows и
					//      пересекает region.cols хотя бы на 1 колонку — добавляем
					//      пересекающийся отрезок в skip.
					for (const auto& bAll : bars) {
						if (!region.ContainsY(bAll.y)) {
							continue;
						}

						const int overlapL = std::max(region.cols.x1, bAll.x1);
						const int overlapR = std::min(region.cols.x2, bAll.x2);

						if (overlapL <= overlapR) {
							auto& vec = skipByRow[bAll.y];
							vec.push_back(Model::SpanX{ overlapL, overlapR });
						}
					}

					// 3.2. Для «детей» дополнительно исключаем их полноценные окна Num/Den
					//      (весь X-коридор ребёнка).
					for (const auto& b : children) {
						// строка самой черты ребёнка
						{
							auto& vec = skipByRow[b.y];
							vec.push_back(Model::SpanX{ b.x1, b.x2 });
						}

						// строки числителя ребёнка
						for (int y = b.numRegion.rows.y1; y <= b.numRegion.rows.y2; ++y) {
							auto& vec = skipByRow[y];
							vec.push_back(Model::SpanX{ b.numRegion.cols.x1, b.numRegion.cols.x2 });
						}

						// строки знаменателя ребёнка
						for (int y = b.denRegion.rows.y1; y <= b.denRegion.rows.y2; ++y) {
							auto& vec = skipByRow[y];
							vec.push_back(Model::SpanX{ b.denRegion.cols.x1, b.denRegion.cols.x2 });
						}
					}

					// 3.3. Сливаем пересекающиеся/соприкасающиеся интервалы по каждой строке.
					for (auto& kv : skipByRow) {
						auto& spans = kv.second;

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

						std::vector<Model::SpanX> merged;

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

					// --- 4) Токенизировать «хвосты» (строго по прямоугольнику region), пропуская skip-интервалы.
					for (int y = region.rows.y1; y <= region.rows.y2; ++y) {
						int x = region.cols.x1;

						const auto it = skipByRow.find(y);
						const bool hasSkips = (it != skipByRow.end());

						const std::vector<Model::SpanX>* skipSpans = nullptr;
						if (hasSkips) {
							skipSpans = &it->second;
						}

						while (x <= region.cols.x2) {
							// Если попали в skip-интервал — перепрыгиваем за него.
							if (hasSkips) {
								bool jumped = false;

								for (const auto& s : *skipSpans) {
									if (x >= s.x1 && x <= s.x2) {
										x = s.x2 + 1;
										jumped = true;
										break;
									}
								}

								if (jumped) {
									continue;
								}
							}

							// Пробелы пропускаем
							while (x <= region.cols.x2 && grid.At(y, x) == ' ') {
								++x;
							}

							if (x > region.cols.x2) {
								break;
							}

							// Если старт токена попадает внутрь skip после пропуска пробелов — перепрыгиваем.
							if (hasSkips) {
								bool jumped = false;

								for (const auto& s : *skipSpans) {
									if (x >= s.x1 && x <= s.x2) {
										x = s.x2 + 1;
										jumped = true;
										break;
									}
								}

								if (jumped) {
									continue;
								}
							}

							// Собираем непробельный «островок» в Symbol
							std::string token{};

							while (x <= region.cols.x2 && grid.At(y, x) != ' ') {
								// Если в процессе наткнулись на skip-интервал — закрываем текущий токен и прыгаем.
								bool inSkip = false;

								if (hasSkips) {
									for (const auto& s : *skipSpans) {
										if (x >= s.x1 && x <= s.x2) {
											inSkip = true;
											break;
										}
									}
								}

								if (inSkip) {
									break;
								}

								token.push_back(grid.At(y, x));
								++x;
							}

							if (!token.empty()) {
								nodes.push_back(
									std::make_unique<Model::Symbol>(
										std::move(token)
									)
								);
							}
						}
					}

					return nodes;
				}
			};
		}
	}
}