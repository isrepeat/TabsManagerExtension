//#pragma once
//#include "FractionBars.h"
//#include <string>
//#include <vector>
//#include <memory>
//
//namespace AsciiMathParser {
//	namespace Core {
//		// ────────────────────────────────────────────────────────────────
//		// ░ INodeVisitor — интерфейс визитора для обхода AST
//		// ────────────────────────────────────────────────────────────────
//		//
//		// Позволяет отделить «как выводим» от «что распознали».
//		// Новые форматы (LaTeX, Markdown, WPF и т.д.) добавляются
//		// как отдельные реализации визитора.
//		//
//		struct INodeVisitor {
//			virtual ~INodeVisitor() = default;
//
//			virtual void Visit(
//				const FractionNode& node
//			) = 0;
//
//			virtual void Visit(
//				const TextRunNode& node
//			) = 0;
//
//			virtual void Visit(
//				const NodeSequence& node
//			) = 0;
//		};
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ IMathNode — базовый интерфейс узла AST
//		// ────────────────────────────────────────────────────────────────
//		//
//		// Любой узел AST обязан уметь «принять» визитора.
//		//
//		struct IMathNode {
//			virtual ~IMathNode() = default;
//
//			virtual void Accept(
//				INodeVisitor& visitor
//			) const = 0;
//		};
//
//		using NodePtr = std::shared_ptr<IMathNode>;
//
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ TextRunNode — «кусок текста»
//		// ────────────────────────────────────────────────────────────────
//		//
//		// На перспективу: будет использоваться для сырых текстовых
//		// фрагментов внутри регионов (NUM/DEN) между структурными
//		// узлами (дробями, суммами и т.д.).
//		//
//		struct TextRunNode : IMathNode {
//			std::string text;
//
//			explicit TextRunNode(
//				std::string text
//			);
//
//			void Accept(
//				INodeVisitor& visitor
//			) const override;
//		};
//
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ NodeSequence — последовательность узлов
//		// ────────────────────────────────────────────────────────────────
//		//
//		// Нужен для случаев, когда в регионе идёт смесь:
//		// текст, операторы, дроби, индексы… Всё это — упорядоченная
//		// последовательность детей.
//		//
//		struct NodeSequence : IMathNode {
//			std::vector<NodePtr> children;
//
//			void Accept(
//				INodeVisitor& visitor
//			) const override;
//		};
//
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ FractionNode — узел-дробь
//		// ────────────────────────────────────────────────────────────────
//		//
//		// Содержит:
//		//  - Собственную «черту» (FracBar) — геометрический якорь,
//		//  - Прямоугольники числителя и знаменателя (Region),
//		//  - Дочерние дроби (вложенные) — найденные внутри NUM/DEN.
//		//
//		struct FractionNode : IMathNode {
//			FracBar bar;                       // геометрия самой черты
//			Model::Region numeratorRegion;     // область числителя
//			Model::Region denominatorRegion;   // область знаменателя
//			std::vector<FractionNode> children; // вложенные дроби
//
//			void Accept(
//				INodeVisitor& visitor
//			) const override;
//		};
//
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ BarTreeBuilder — строитель дерева дробей из списка FracBar
//		// ────────────────────────────────────────────────────────────────
//		//
//		// Алгоритм:
//		//  1) Для каждой черты вычисляет NUM/DEN регионы (через FractionBars).
//		//  2) Для каждой черты подбирает «родителя» (в чей NUM или DEN она попадает).
//		//     Критерий — попадание по Y + достаточное перекрытие по X (эвристика).
//		//  3) По связям «родитель→дети» строит лес корней FractionNode.
//		//  4) Умеет извлекать «видимый» текст из регионов с маскированием
//		//     дочерних областей (чтобы не дублировать).
//		//
//		class BarTreeBuilder {
//		public:
//			explicit BarTreeBuilder(
//				const Model::AsciiGrid& asciiGrid
//			);
//
//			// Строит лес корневых дробей из найденных горизонтальных черт.
//			std::vector<FractionNode> Build(
//				const std::vector<FracBar>& bars
//			) const;
//
//			// Отладочный дамп дерева (кратко по каждой дроби).
//			std::string DumpTree(
//				const std::vector<FractionNode>& roots
//			) const;
//
//			// Простейший вывод в стиле LaTeX: \frac{ ... }{ ... }
//			// (без полноценной токенизации текста внутри NUM/DEN).
//			std::string RenderPseudoLatex(
//				const std::vector<FractionNode>& roots
//			) const;
//
//		private:
//			const Model::AsciiGrid& grid;
//
//			// Проверка попадания по вертикали: y ∈ region.rows
//			bool IsInsideY(
//				int y,
//				const Model::Region& region
//			) const;
//
//			// Длина перекрытия по X (для оценки «насколько бар попал внутрь»).
//			int OverlapX(
//				const Model::Region& a,
//				const Model::Region& b
//			) const;
//
//			// Подбор индекса родителя для i-й черты.
//			// Возвращает -1, если родителя нет (значит это корень).
//			int ChooseParentIndex(
//				const std::vector<FracBar>& bars,
//				const std::vector<Model::Region>& nums,
//				const std::vector<Model::Region>& dens,
//				int childIndex
//			) const;
//
//			// Извлекает строковое содержимое региона, «маскируя» (пропуская)
//			// прямоугольники holes (черты+их NUM/DEN), чтобы убрать вложенное.
//			std::string RectTextMasked(
//				const Model::Region& region,
//				const std::vector<Model::Region>& holes
//			) const;
//
//			// Рекурсивный дамп одного узла.
//			void DumpNode(
//				const FractionNode& node,
//				int level,
//				std::string& out
//			) const;
//
//			// Рекурсивный LaTeX-подобный рендер одного узла.
//			std::string RenderNode(
//				const FractionNode& node
//			) const;
//		};
//
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ LatexRenderer — пример визитора (прототип)
//		// ────────────────────────────────────────────────────────────────
//		//
//		// Демонстрация того, как можно отделить вывод от данных.
//		// Сейчас использует заглушки для FractionNode (чтобы не дублировать
//		// RenderNode внутри визитора). После появления полноценной токенизации
//		// NUM/DEN (TextRun/NodeSequence) будет собирать LaTeX честно из детей.
//		//
//		class LatexRenderer : public INodeVisitor {
//		public:
//			LatexRenderer();
//
//			void Visit(
//				const FractionNode& node
//			) override;
//
//			void Visit(
//				const TextRunNode& node
//			) override;
//
//			void Visit(
//				const NodeSequence& node
//			) override;
//
//			std::string Str() const;
//
//		private:
//			std::string buffer;
//		};
//
//
//		// ────────────────────────────────────────────────────────────────
//		// ░ DebugTreeDumper — визитор для отладочного текстового дампа
//		// ────────────────────────────────────────────────────────────────
//		class DebugTreeDumper : public INodeVisitor {
//		public:
//			DebugTreeDumper();
//
//			void Visit(
//				const FractionNode& node
//			) override;
//
//			void Visit(
//				const TextRunNode& node
//			) override;
//
//			void Visit(
//				const NodeSequence& node
//			) override;
//
//			std::string Str() const;
//
//		private:
//			int level;
//			std::string buffer;
//		};
//
//	}
//}