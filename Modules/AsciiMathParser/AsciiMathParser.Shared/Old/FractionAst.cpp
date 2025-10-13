//#include "FractionAst.h"
//#include <algorithm>
//#include <sstream>
//#include <span>
//
//namespace AsciiMathParser {
//	namespace Core {
//		// ────────────────────────────────────────────────────────────────
//			// ░ TextRunNode
//			// ────────────────────────────────────────────────────────────────
//
//		TextRunNode::TextRunNode(
//			std::string text
//		)
//			: text{ std::move(text) } {
//		}
//
//		void TextRunNode::Accept(
//			INodeVisitor& visitor
//		) const {
//			visitor.Visit(*this);
//		}
//
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ NodeSequence
//		// ────────────────────────────────────────────────────────────────
//
//		void NodeSequence::Accept(
//			INodeVisitor& visitor
//		) const {
//			visitor.Visit(*this);
//		}
//
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ FractionNode
//		// ────────────────────────────────────────────────────────────────
//
//		void FractionNode::Accept(
//			INodeVisitor& visitor
//		) const {
//			visitor.Visit(*this);
//		}
//
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ BarTreeBuilder — реализация
//		// ────────────────────────────────────────────────────────────────
//
//		BarTreeBuilder::BarTreeBuilder(
//			const Model::AsciiGrid& asciiGrid
//		)
//			: grid{ asciiGrid } {
//		}
//
//		bool BarTreeBuilder::IsInsideY(
//			int y,
//			const Model::Region& region
//		) const {
//			return region.ContainsY(y);
//		}
//
//		int BarTreeBuilder::OverlapX(
//			const Model::Region& a,
//			const Model::Region& b
//		) const {
//			return a.OverlapX(b);
//		}
//
//		// Подбор индекса родителя для дочерней черты:
//		//  - Родительский NUM/DEN должен содержать по Y строку дочерней черты,
//		//  - Перекрытие по X должно быть «достаточным» (эвристика τ≈0.7),
//		//  - При равенстве критериев — выбираем более «широкого» родителя,
//		//    затем — более близкого по |Δy|, затем — с меньшим y (стабильность).
//		int BarTreeBuilder::ChooseParentIndex(
//			const std::vector<FracBar>& bars,
//			const std::vector<Model::Region>& nums,
//			const std::vector<Model::Region>& dens,
//			int childIndex
//		) const {
//			const FracBar& childBar = bars[static_cast<std::size_t>(childIndex)];
//			const Model::Region childBand = childBar.AsBandRegion();
//
//			int bestIndex = -1;
//			int bestOverlap = -1;
//			int bestWidth = -1;
//			int bestYDistance = 1'000'000;
//			int bestY = 1'000'000;
//
//			for (int i = 0; i < static_cast<int>(bars.size()); ++i) {
//				if (i == childIndex) {
//					continue;
//				}
//
//				const Model::Region& numR = nums[static_cast<std::size_t>(i)];
//				const Model::Region& denR = dens[static_cast<std::size_t>(i)];
//
//				const bool insideNum = this->IsInsideY(childBar.y, numR);
//				const bool insideDen = this->IsInsideY(childBar.y, denR);
//				if (!insideNum && !insideDen) {
//					continue;
//				}
//
//				// Суммарное перекрытие по X с обеими областями родителя.
//				const int ox =
//					this->OverlapX(numR, childBand) +
//					this->OverlapX(denR, childBand);
//
//				const int childWidth = childBand.Width();
//				const int parentWidth = bars[static_cast<std::size_t>(i)].AsBandRegion().Width();
//
//				// Эвристика «достаточности» перекрытия: ox / min(parent, child) ≥ 0.7
//				const int minWidth = std::min(parentWidth, childWidth);
//				const bool enough = (ox * 10 >= minWidth * 7);
//
//				if (!enough) {
//					continue;
//				}
//
//				const int yDist = std::abs(bars[static_cast<std::size_t>(i)].y - childBar.y);
//
//				bool better = false;
//				if (ox > bestOverlap) {
//					better = true;
//				}
//				else if (ox == bestOverlap) {
//					if (parentWidth > bestWidth) {
//						better = true;
//					}
//					else if (parentWidth == bestWidth) {
//						if (yDist < bestYDistance) {
//							better = true;
//						}
//						else if (yDist == bestYDistance) {
//							if (bars[static_cast<std::size_t>(i)].y < bestY) {
//								better = true;
//							}
//						}
//					}
//				}
//
//				if (better) {
//					bestIndex = i;
//					bestOverlap = ox;
//					bestWidth = parentWidth;
//					bestYDistance = yDist;
//					bestY = bars[static_cast<std::size_t>(i)].y;
//				}
//			}
//
//			return bestIndex;
//		}
//
//		std::vector<FractionNode> BarTreeBuilder::Build(
//			const std::vector<FracBar>& bars
//		) const {
//			if (bars.empty()) {
//				return {};
//			}
//
//			// 1) Считаем регионы NUM/DEN для каждой найденной черты.
//			FractionBars helper{ this->grid };
//
//			std::vector<Model::Region> nums{};
//			std::vector<Model::Region> dens{};
//			nums.resize(bars.size());
//			dens.resize(bars.size());
//
//			for (std::size_t i = 0; i < bars.size(); ++i) {
//				helper.ComputeNumDenRects(
//					bars[i],
//					nums[i],
//					dens[i]
//				);
//			}
//
//			// 2) Для каждой черты находим потенциального родителя.
//			std::vector<int> parentIdx{};
//			parentIdx.resize(bars.size(), -1);
//
//			for (int i = 0; i < static_cast<int>(bars.size()); ++i) {
//				parentIdx[static_cast<std::size_t>(i)] = this->ChooseParentIndex(
//					bars,
//					nums,
//					dens,
//					i
//				);
//			}
//
//			// 3) Инвертируем связи: список детей у каждого потенциального родителя.
//			std::vector<std::vector<int>> children{};
//			children.resize(bars.size());
//
//			for (int i = 0; i < static_cast<int>(bars.size()); ++i) {
//				const int p = parentIdx[static_cast<std::size_t>(i)];
//				if (p >= 0) {
//					children[static_cast<std::size_t>(p)].push_back(i);
//				}
//			}
//
//			// 4) Собираем корни (те, у кого нет родителя), и рекурсивно добавляем потомков.
//			std::vector<FractionNode> roots{};
//
//			for (int i = 0; i < static_cast<int>(bars.size()); ++i) {
//				if (parentIdx[static_cast<std::size_t>(i)] < 0) {
//					FractionNode root{
//						bars[static_cast<std::size_t>(i)],
//						nums[static_cast<std::size_t>(i)],
//						dens[static_cast<std::size_t>(i)],
//						{}
//					};
//
//					// Ранжируем детей слева-направо (и немного по y для стабильности).
//					std::vector<int> myKids = children[static_cast<std::size_t>(i)];
//					std::sort(
//						myKids.begin(),
//						myKids.end(),
//						[&bars](
//							int a,
//							int b
//							) {
//								if (bars[static_cast<std::size_t>(a)].x1 != bars[static_cast<std::size_t>(b)].x1) {
//									return bars[static_cast<std::size_t>(a)].x1 < bars[static_cast<std::size_t>(b)].x1;
//								}
//								return bars[static_cast<std::size_t>(a)].y < bars[static_cast<std::size_t>(b)].y;
//						}
//					);
//
//					for (int kid : myKids) {
//						FractionNode child{
//							bars[static_cast<std::size_t>(kid)],
//							nums[static_cast<std::size_t>(kid)],
//							dens[static_cast<std::size_t>(kid)],
//							{}
//						};
//
//						for (int grand : children[static_cast<std::size_t>(kid)]) {
//							FractionNode grandNode{
//								bars[static_cast<std::size_t>(grand)],
//								nums[static_cast<std::size_t>(grand)],
//								dens[static_cast<std::size_t>(grand)],
//								{}
//							};
//							child.children.push_back(grandNode);
//						}
//
//						root.children.push_back(child);
//					}
//
//					roots.push_back(root);
//				}
//			}
//
//			return roots;
//		}
//
//		// Извлечение текста из прямоугольника с учётом «масок» (holes).
//		// Маскируем: полосы дочерних дробей и их NUM/DEN, чтобы не склеивать
//		// родительский текст с вложенным.
//		std::string BarTreeBuilder::RectTextMasked(
//			const Model::Region& region,
//			const std::vector<Model::Region>& holes
//		) const {
//			std::string out{};
//
//			for (int y = region.rows.y1; y <= region.rows.y2; ++y) {
//				bool anyNonSpace = false;
//
//				// Собираем интервал(ы) по X, которые надо пропустить на строке y.
//				std::vector<std::pair<int, int>> cut{};
//				for (const auto& h : holes) {
//					if (y >= h.rows.y1 && y <= h.rows.y2) {
//						const int lx = std::max(region.cols.x1, h.cols.x1);
//						const int rx = std::min(region.cols.x2, h.cols.x2);
//						if (lx <= rx) {
//							cut.emplace_back(lx, rx);
//						}
//					}
//				}
//				std::sort(cut.begin(), cut.end());
//
//				int x = region.cols.x1;
//				while (x <= region.cols.x2) {
//					bool skip = false;
//
//					for (const auto& c : cut) {
//						if (x >= c.first && x <= c.second) {
//							x = c.second + 1; // прыжок за маску
//							skip = true;
//							break;
//						}
//					}
//					if (skip) {
//						continue;
//					}
//
//					const char ch = this->grid.At(y, x);
//					if (ch != ' ') {
//						anyNonSpace = true;
//					}
//					out.push_back(ch);
//					++x;
//				}
//
//				// Мини-выравнивание: отделяем строки пробелом,
//				// но только если текущая строка содержала непробельные символы.
//				if (y < region.rows.y2) {
//					if (anyNonSpace) {
//						out.push_back(' ');
//					}
//				}
//			}
//
//			// Обрезаем ведущие/замыкающие пробелы.
//			while (!out.empty() && out.front() == ' ') {
//				out.erase(out.begin());
//			}
//			while (!out.empty() && out.back() == ' ') {
//				out.pop_back();
//			}
//
//			return out;
//		}
//
//		void BarTreeBuilder::DumpNode(
//			const FractionNode& node,
//			int level,
//			std::string& out
//		) const {
//			// Разделяем дочерние дроби на «лежащие в NUM» и «лежащие в DEN».
//			std::vector<const FractionNode*> numKids{};
//			std::vector<const FractionNode*> denKids{};
//
//			for (const auto& ch : node.children) {
//				if (this->IsInsideY(ch.bar.y, node.numeratorRegion)) {
//					numKids.push_back(&ch);
//				}
//				else if (this->IsInsideY(ch.bar.y, node.denominatorRegion)) {
//					denKids.push_back(&ch);
//				}
//			}
//
//			// Коллекция «масок» — зоны, которые нужно исключить при выборе текста.
//			auto collectHoles = [](
//				const std::vector<const FractionNode*>& kids
//				) -> std::vector<Model::Region> {
//					std::vector<Model::Region> holes{};
//					holes.reserve(kids.size() * 3);
//					for (auto* k : kids) {
//						holes.push_back(k->bar.AsBandRegion());
//						holes.push_back(k->numeratorRegion);
//						holes.push_back(k->denominatorRegion);
//					}
//					return holes;
//				};
//
//			const std::vector<Model::Region> numHoles = collectHoles(numKids);
//			const std::vector<Model::Region> denHoles = collectHoles(denKids);
//
//			const std::string numText = this->RectTextMasked(
//				node.numeratorRegion,
//				numHoles
//			);
//			const std::string denText = this->RectTextMasked(
//				node.denominatorRegion,
//				denHoles
//			);
//
//			std::ostringstream line{};
//			line
//				<< std::string(static_cast<std::size_t>(level * 2), ' ')
//				<< "Frac[y=" << node.bar.y
//				<< ", x=(" << node.bar.x1 << ".." << node.bar.x2 << ")]  "
//				<< "NUM:\"" << numText << "\"  "
//				<< "DEN:\"" << denText << "\"\n";
//
//			out += line.str();
//
//			for (const auto& ch : node.children) {
//				this->DumpNode(
//					ch,
//					level + 1,
//					out
//				);
//			}
//		}
//
//		std::string BarTreeBuilder::DumpTree(
//			const std::vector<FractionNode>& roots
//		) const {
//			std::string out{};
//			for (const auto& r : roots) {
//				this->DumpNode(
//					r,
//					0,
//					out
//				);
//			}
//			return out;
//		}
//
//		std::string BarTreeBuilder::RenderNode(
//			const FractionNode& node
//		) const {
//			// Готовим списки вложенных дробей по NUM и DEN.
//			std::vector<const FractionNode*> numKids{};
//			std::vector<const FractionNode*> denKids{};
//
//			for (const auto& ch : node.children) {
//				if (this->IsInsideY(ch.bar.y, node.numeratorRegion)) {
//					numKids.push_back(&ch);
//				}
//				else if (this->IsInsideY(ch.bar.y, node.denominatorRegion)) {
//					denKids.push_back(&ch);
//				}
//			}
//
//			auto collectHoles = [](
//				const std::vector<const FractionNode*>& kids
//				) -> std::vector<Model::Region> {
//					std::vector<Model::Region> holes{};
//					holes.reserve(kids.size() * 3);
//					for (auto* k : kids) {
//						holes.push_back(k->bar.AsBandRegion());
//						holes.push_back(k->numeratorRegion);
//						holes.push_back(k->denominatorRegion);
//					}
//					return holes;
//				};
//
//			const std::vector<Model::Region> numHoles = collectHoles(numKids);
//			const std::vector<Model::Region> denHoles = collectHoles(denKids);
//
//			std::ostringstream num{};
//			std::ostringstream den{};
//
//			// Базовый текст региона без вложенных областей.
//			num << this->RectTextMasked(
//				node.numeratorRegion,
//				numHoles
//			);
//			den << this->RectTextMasked(
//				node.denominatorRegion,
//				denHoles
//			);
//
//			// После текста — добавляем рендер вложенных дробей.
//			for (auto* k : numKids) {
//				num << " " << this->RenderNode(*k);
//			}
//			for (auto* k : denKids) {
//				den << " " << this->RenderNode(*k);
//			}
//
//			std::ostringstream oss{};
//			oss
//				<< "\\frac{ "
//				<< num.str()
//				<< " }{ "
//				<< den.str()
//				<< " }";
//
//			return oss.str();
//		}
//
//		std::string BarTreeBuilder::RenderPseudoLatex(
//			const std::vector<FractionNode>& roots
//		) const {
//			std::ostringstream oss{};
//			bool first = true;
//
//			for (const auto& r : roots) {
//				if (!first) {
//					oss << "  ";
//				}
//				else {
//					first = false;
//				}
//				oss << this->RenderNode(r);
//			}
//
//			return oss.str();
//		}
//
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ LatexRenderer — прототип визитора
//		// ────────────────────────────────────────────────────────────────
//
//		LatexRenderer::LatexRenderer()
//			: buffer{} {
//		}
//
//		void LatexRenderer::Visit(
//			const FractionNode& /*node*/
//		) {
//			// Здесь можно задействовать BarTreeBuilder::RenderNode,
//			// но для чистой архитектуры визитор обычно работает по AST
//			// (после того как NUM/DEN станут NodeSequence/TextRun).
//			// Пока оставляем заглушку:
//			this->buffer += "\\frac{...}{...}";
//		}
//
//		void LatexRenderer::Visit(
//			const TextRunNode& node
//		) {
//			this->buffer += node.text;
//		}
//
//		void LatexRenderer::Visit(
//			const NodeSequence& node
//		) {
//			for (std::size_t i = 0; i < node.children.size(); ++i) {
//				node.children[i]->Accept(*this);
//				if (i + 1 < node.children.size()) {
//					this->buffer += " ";
//				}
//			}
//		}
//
//		std::string LatexRenderer::Str() const {
//			return this->buffer;
//		}
//
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ DebugTreeDumper — визитор дампа
//		// ────────────────────────────────────────────────────────────────
//
//		DebugTreeDumper::DebugTreeDumper()
//			: level{ 0 }
//			, buffer{} {
//		}
//
//		void DebugTreeDumper::Visit(
//			const FractionNode& node
//		) {
//			this->buffer += std::string(static_cast<std::size_t>(this->level * 2), ' ');
//			this->buffer += "Frac[y=" + std::to_string(node.bar.y)
//				+ ", x=(" + std::to_string(node.bar.x1) + ".." + std::to_string(node.bar.x2) + ")]\n";
//
//			++this->level;
//			for (const auto& ch : node.children) {
//				ch.Accept(*this);
//			}
//			--this->level;
//		}
//
//		void DebugTreeDumper::Visit(
//			const TextRunNode& node
//		) {
//			this->buffer += std::string(static_cast<std::size_t>(this->level * 2), ' ');
//			this->buffer += "Text: \"" + node.text + "\"\n";
//		}
//
//		void DebugTreeDumper::Visit(
//			const NodeSequence& node
//		) {
//			this->buffer += std::string(static_cast<std::size_t>(this->level * 2), ' ');
//			this->buffer += "Seq:\n";
//			++this->level;
//			for (const auto& c : node.children) {
//				c->Accept(*this);
//			}
//			--this->level;
//		}
//
//		std::string DebugTreeDumper::Str() const {
//			return this->buffer;
//		}
//
//	}
//}