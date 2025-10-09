#pragma once
#include "FractionBars.h"
#include <string>
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		struct FractionNode {
			FracBar bar;
			Rect numeratorRect;
			Rect denominatorRect;
			std::vector<FractionNode> children;
		};

		class BarTreeBuilder {
		public:
			explicit BarTreeBuilder(
				const AsciiGrid& asciiGrid
			);

			// Построить лес верхнеуровневых дробей (обычно один корень)
			std::vector<FractionNode> Build(
				const std::vector<FracBar>& bars
			) const;

			//int DebugParentIndex(int child) const;

			// Быстрый вывод дерева (для отладки/демо)
			std::string DumpTree(
				const std::vector<FractionNode>& roots
			) const;

			// Черновой рендер в псевдо-LaTeX (только \frac, без Σ/∏ и индексов)
			std::string RenderPseudoLatex(
				const std::vector<FractionNode>& roots
			) const;

		private:
			const AsciiGrid& grid;
			// кэш последнего вызова Build()
			std::vector<FracBar> cachedBars;
			std::vector<Rect> cachedNums;
			std::vector<Rect> cachedDens;

			bool IsInsideY(
				int y,
				const Rect& rect
			) const;

			int OverlapX(
				const Rect& a,
				const Rect& b
			) const;

			// Родительство по правилам из объяснения
			int ChooseParentIndex(
				const std::vector<FracBar>& bars,
				const std::vector<Rect>& nums,
				const std::vector<Rect>& dens,
				int childIndex
			) const;

			// Собрать узел FractionNode из bar + rects и прикрепить детей
			void AttachChildren(
				FractionNode& parentNode,
				const std::vector<int>& childIndices,
				const std::vector<FracBar>& bars,
				const std::vector<Rect>& nums,
				const std::vector<Rect>& dens
			) const;

			// Новый: вернёт текст r, но "с дырками" по прямоугольникам маски
			std::string RectTextMasked(
				const Rect& r,
				const std::vector<Rect>& holes
			) const;

			// Рекурсивный рендер \frac{..}{..}
			std::string RenderNode(
				const FractionNode& node
			) const;

			// Рекурсивный дамп
			void DumpNode(
				const FractionNode& node,
				int level,
				std::string& out
			) const;
		};
	}
}