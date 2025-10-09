#include "FractionAst.h"
#include <algorithm>
#include <sstream>
#include <span>

namespace AsciiMathParser {
	namespace Core {
		BarTreeBuilder::BarTreeBuilder(
			const AsciiGrid& asciiGrid
		)
			: grid{ asciiGrid } {
			std::span<int> x;
		}

		bool BarTreeBuilder::IsInsideY(
			int y,
			const Rect& rect
		) const {
			if (y < rect.y1) {
				return false;
			}
			if (y > rect.y2) {
				return false;
			}
			return true;
		}

		int BarTreeBuilder::OverlapX(
			const Rect& a,
			const Rect& b
		) const {
			const int left = std::max(a.x1, b.x1);
			const int right = std::min(a.x2, b.x2);
			if (right < left) {
				return 0;
			}
			return (right - left + 1);
		}

		int BarTreeBuilder::ChooseParentIndex(
			const std::vector<FracBar>& bars,
			const std::vector<Rect>& nums,
			const std::vector<Rect>& dens,
			int childIndex
		) const {
			const FracBar& childBar = bars[static_cast<std::size_t>(childIndex)];
			const Rect childRect{
				childBar.y,
				childBar.y,
				childBar.x1,
				childBar.x2
			};

			int bestIndex = -1;
			int bestOverlap = -1;
			int bestWidth = -1;
			int bestYDistance = 1'000'000;
			int bestY = 1'000'000;

			for (int i = 0; i < static_cast<int>(bars.size()); ++i) {
				if (i == childIndex) {
					continue;
				}

				const Rect& numR = nums[static_cast<std::size_t>(i)];
				const Rect& denR = dens[static_cast<std::size_t>(i)];

				const bool insideNum = this->IsInsideY(childBar.y, numR);
				const bool insideDen = this->IsInsideY(childBar.y, denR);

				if (!insideNum && !insideDen) {
					continue;
				}

				const int ox = this->OverlapX(numR, childRect) + this->OverlapX(denR, childRect);
				const int childWidth = (childBar.x2 - childBar.x1 + 1);

				// Порог существенного перекрытия по X (от min ширины)
				const int parentWidth = (bars[static_cast<std::size_t>(i)].x2 - bars[static_cast<std::size_t>(i)].x1 + 1);
				const int minWidth = std::min(parentWidth, childWidth);
				const bool enough = (ox * 10 >= minWidth * 7); // τx ≈ 0.7

				if (!enough) {
					continue;
				}

				const int yDist = std::abs(bars[static_cast<std::size_t>(i)].y - childBar.y);

				bool better = false;
				if (ox > bestOverlap) {
					better = true;
				}
				else if (ox == bestOverlap) {
					if (parentWidth > bestWidth) {
						better = true;
					}
					else if (parentWidth == bestWidth) {
						if (yDist < bestYDistance) {
							better = true;
						}
						else if (yDist == bestYDistance) {
							if (bars[static_cast<std::size_t>(i)].y < bestY) {
								better = true;
							}
						}
					}
				}

				if (better) {
					bestIndex = i;
					bestOverlap = ox;
					bestWidth = parentWidth;
					bestYDistance = yDist;
					bestY = bars[static_cast<std::size_t>(i)].y;
				}
			}

			return bestIndex;
		}

		void BarTreeBuilder::AttachChildren(
			FractionNode& parentNode,
			const std::vector<int>& childIndices,
			const std::vector<FracBar>& bars,
			const std::vector<Rect>& nums,
			const std::vector<Rect>& dens
		) const {
			for (int idx : childIndices) {
				FractionNode child{
					bars[static_cast<std::size_t>(idx)],
					nums[static_cast<std::size_t>(idx)],
					dens[static_cast<std::size_t>(idx)],
					{}
				};
				this->AttachChildren(
					child,
					{},
					bars,
					nums,
					dens
				);
				parentNode.children.push_back(child);
			}
		}

		std::vector<FractionNode> BarTreeBuilder::Build(
			const std::vector<FracBar>& bars
		) const {
			if (bars.empty()) {
				return {};
			}

			FractionBars helper{ this->grid };

			std::vector<Rect> nums{};
			std::vector<Rect> dens{};
			nums.resize(bars.size());
			dens.resize(bars.size());

			for (std::size_t i = 0; i < bars.size(); ++i) {
				helper.ComputeNumDenRects(
					bars[i],
					nums[i],
					dens[i]
				);
			}

			// Родители
			std::vector<int> parentIdx{};
			parentIdx.resize(bars.size(), -1);
			for (int i = 0; i < static_cast<int>(bars.size()); ++i) {
				parentIdx[static_cast<std::size_t>(i)] = this->ChooseParentIndex(
					bars,
					nums,
					dens,
					i
				);
			}

			// Дети каждого узла
			std::vector<std::vector<int>> children{};
			children.resize(bars.size());
			for (int i = 0; i < static_cast<int>(bars.size()); ++i) {
				const int p = parentIdx[static_cast<std::size_t>(i)];
				if (p >= 0) {
					children[static_cast<std::size_t>(p)].push_back(i);
				}
			}

			// Сбор корней
			std::vector<FractionNode> roots{};
			for (int i = 0; i < static_cast<int>(bars.size()); ++i) {
				if (parentIdx[static_cast<std::size_t>(i)] < 0) {
					FractionNode root{
						bars[static_cast<std::size_t>(i)],
						nums[static_cast<std::size_t>(i)],
						dens[static_cast<std::size_t>(i)],
						{}
					};

					// Рекурсивное прикрепление детей
					std::vector<int> myKids = children[static_cast<std::size_t>(i)];

					// Отсортируем детей слева-направо (по x1)
					std::sort(
						myKids.begin(),
						myKids.end(),
						[&bars](
							int a,
							int b
							) {
								if (bars[static_cast<std::size_t>(a)].x1 != bars[static_cast<std::size_t>(b)].x1) {
									return bars[static_cast<std::size_t>(a)].x1 < bars[static_cast<std::size_t>(b)].x1;
								}
								return bars[static_cast<std::size_t>(a)].y < bars[static_cast<std::size_t>(b)].y;
						}
					);

					for (int kid : myKids) {
						// Вложенные дети самого child мы добавим при проходе вниз
						// (проще — соберём node полностью рекурсивно)
						// Здесь создадим заготовку, а потом углубимся:
						// Но чтобы не усложнять — просто вызовем Build заново ограниченно нельзя,
						// поэтому сформируем child узел напрямую:

						FractionNode child{
							bars[static_cast<std::size_t>(kid)],
							nums[static_cast<std::size_t>(kid)],
							dens[static_cast<std::size_t>(kid)],
							{}
						};

						// Прикрепим внуков:
						for (int grand : children[static_cast<std::size_t>(kid)]) {
							FractionNode grandNode{
								bars[static_cast<std::size_t>(grand)],
								nums[static_cast<std::size_t>(grand)],
								dens[static_cast<std::size_t>(grand)],
								{}
							};
							child.children.push_back(grandNode);
						}

						root.children.push_back(child);
					}

					roots.push_back(root);
				}
			}

			return roots;
		}


		std::string BarTreeBuilder::RectTextMasked(
			const Rect& r,
			const std::vector<Rect>& holes
		) const {
			std::string out{};
			for (int y = r.y1; y <= r.y2; ++y) {
				bool any = false;

				// Для строки y вычислим интервалы, которые нужно скрыть
				// (пересечение holes с текущей строкой)
				std::vector<std::pair<int, int>> cut{};
				for (const auto& h : holes) {
					if (y >= h.y1 && y <= h.y2) {
						const int lx = std::max(r.x1, h.x1);
						const int rx = std::min(r.x2, h.x2);
						if (lx <= rx) {
							cut.emplace_back(lx, rx);
						}
					}
				}
				std::sort(
					cut.begin(),
					cut.end()
				);

				// Пройдём по сегментам строки r.x1..r.x2, пропуская вырезанные интервалы
				int x = r.x1;
				while (x <= r.x2) {
					// если x попадает в дырку — перескочим
					bool skip = false;
					for (const auto& c : cut) {
						if (x >= c.first && x <= c.second) {
							x = c.second + 1;
							skip = true;
							break;
						}
					}
					if (skip) {
						continue;
					}

					// обычный символ
					const char ch = this->grid.At(y, x);
					if (ch != ' ') {
						any = true;
					}
					out.push_back(ch);
					++x;
				}

				if (y < r.y2) {
					if (any) {
						out.push_back(' ');
					}
				}
			}

			// trim
			while (!out.empty() && out.front() == ' ') {
				out.erase(out.begin());
			}
			while (!out.empty() && out.back() == ' ') {
				out.pop_back();
			}
			return out;
		}


		void BarTreeBuilder::DumpNode(
			const FractionNode& node,
			int level,
			std::string& out
		) const {
			// ─────────────────────────────────────────────
			// 1. Разделяем детей по принадлежности к NUM / DEN
			// ─────────────────────────────────────────────
			std::vector<const FractionNode*> numKids{};
			std::vector<const FractionNode*> denKids{};

			for (const auto& ch : node.children) {
				if (this->IsInsideY(ch.bar.y, node.numeratorRect)) {
					numKids.push_back(&ch);
				}
				else if (this->IsInsideY(ch.bar.y, node.denominatorRect)) {
					denKids.push_back(&ch);
				}
			}

			auto collectHoles = [](
				const std::vector<const FractionNode*>& kids
				) -> std::vector<Rect> {
					std::vector<Rect> holes{};
					holes.reserve(kids.size() * 3);
					for (auto* k : kids) {
						holes.push_back(Rect{ k->bar.y, k->bar.y, k->bar.x1, k->bar.x2 });
						holes.push_back(k->numeratorRect);
						holes.push_back(k->denominatorRect);
					}
					return holes;
				};

			const std::vector<Rect> numHoles = collectHoles(numKids);
			const std::vector<Rect> denHoles = collectHoles(denKids);

			// ─────────────────────────────────────────────
			// 2. Извлекаем текст с учётом масок (дыр)
			// ─────────────────────────────────────────────
			std::string numText = this->RectTextMasked(
				node.numeratorRect,
				numHoles
			);
			std::string denText = this->RectTextMasked(
				node.denominatorRect,
				denHoles
			);

			// ─────────────────────────────────────────────
			// 3. Формируем отладочную строку
			// ─────────────────────────────────────────────
			std::ostringstream line{};
			line
				<< std::string(static_cast<std::size_t>(level * 2), ' ')
				<< "Frac[y=" << node.bar.y
				<< ", x=(" << node.bar.x1 << ".." << node.bar.x2 << ")]  "
				<< "NUM:\"" << numText << "\"  "
				<< "DEN:\"" << denText << "\"\n";

			out += line.str();

			// ─────────────────────────────────────────────
			// 4. Рекурсивный обход потомков
			// ─────────────────────────────────────────────
			for (const auto& ch : node.children) {
				this->DumpNode(
					ch,
					level + 1,
					out
				);
			}
		}

		std::string BarTreeBuilder::DumpTree(
			const std::vector<FractionNode>& roots
		) const {
			std::string out{};
			for (const auto& r : roots) {
				this->DumpNode(
					r,
					0,
					out
				);
			}
			return out;
		}


		std::string BarTreeBuilder::RenderNode(
			const FractionNode& node
		) const {
			// Разделим детей по тому, попадает ли их полоса в NUM или в DEN родителя
			std::vector<const FractionNode*> numKids{};
			std::vector<const FractionNode*> denKids{};

			for (const auto& ch : node.children) {
				if (this->IsInsideY(ch.bar.y, node.numeratorRect)) {
					numKids.push_back(&ch);
				}
				else if (this->IsInsideY(ch.bar.y, node.denominatorRect)) {
					denKids.push_back(&ch);
				}
			}

			// Построим "дырки": вырезаем у родителя все области детей
			auto collectHoles = [](
				const std::vector<const FractionNode*>& kids
				) -> std::vector<Rect> {
					std::vector<Rect> holes{};
					holes.reserve(kids.size() * 3);
					for (auto* k : kids) {
						// Вырежем и полосу, и прямоугольники num/den ребёнка
						holes.push_back(Rect{ k->bar.y, k->bar.y, k->bar.x1, k->bar.x2 });
						holes.push_back(k->numeratorRect);
						holes.push_back(k->denominatorRect);
					}
					return holes;
				};

			const std::vector<Rect> numHoles = collectHoles(numKids);
			const std::vector<Rect> denHoles = collectHoles(denKids);

			std::ostringstream num{};
			std::ostringstream den{};

			// Текст без дочерних областей
			num << this->RectTextMasked(
				node.numeratorRect,
				numHoles
			);
			den << this->RectTextMasked(
				node.denominatorRect,
				denHoles
			);

			// Добавим детей (пока просто в хвост; при желании можно упорядочить по x и вставлять маркерами)
			for (auto* k : numKids) {
				num << " " << this->RenderNode(*k);
			}
			for (auto* k : denKids) {
				den << " " << this->RenderNode(*k);
			}

			std::ostringstream oss{};
			oss
				<< "\\frac{ "
				<< num.str()
				<< " }{ "
				<< den.str()
				<< " }";

			return oss.str();
		}

		std::string BarTreeBuilder::RenderPseudoLatex(
			const std::vector<FractionNode>& roots
		) const {
			std::ostringstream oss{};
			bool first = true;
			for (const auto& r : roots) {
				if (!first) {
					oss << "  ";
				}
				else {
					first = false;
				}
				oss << this->RenderNode(r);
			}
			return oss.str();
		}
	}
}