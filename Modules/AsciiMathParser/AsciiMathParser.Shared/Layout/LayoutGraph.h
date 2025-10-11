#pragma once
#include "../Model/Grid.h"
#include <cstdint>
#include <string>
#include <vector>

//
// ░ Description
// ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
// 
// LayoutGraph — это не «граф» в математическом смысле, а обобщённая структура,
// описывающая пространственные отношения элементов формулы на ASCII-сетке.
//
// Вместо того чтобы работать с абсолютными координатами x/y, LayoutGraph 
// абстрактные *узлы* (элементы формулы) и *рёбра* (отношения между ними):
// 
// ░ Узлы (LayoutNode):
//     Каждый узел представляет собой отдельный элемент на ASCII-плоскости:
//       • текстовый сегмент ("text-run")
//       • символ оператора ("+","=","Σ","√" и т.д.)
//       • графический элемент ("bar" — черта дроби, скобка и т.п.)
//
//     У каждого узла есть Region (координаты на сетке) и роль (role),
//     определяющая тип содержимого.
//
// ░ Рёбра (LayoutEdge):
//     Каждое ребро выражает *пространственное отношение* между узлами.
//     Поддерживаются типы:
//       • Above   — один элемент выше другого
//       • Below   — ниже
//       • LeftOf  — левее
//       • RightOf — правее
//       • Inside  — полностью внутри
//       • Adjacent — примыкает (соседство)
//
//     Эти отношения не зависят от конкретного смысла символов;
//     они описывают чисто топологическое расположение.
//
// Идея в том, что любой математический элемент можно описать через взаимное положение:
//     «кто над кем и где находится»
//
// Например:
//     a + b
//    -------   →  создаёт узел "bar" и рёбра:
//       z           ("a+b") --Above--> ("bar")
//                   ("bar") --Above--> ("z")
//
// Такая топологическая модель делает систему расширяемой:
//     новая фича (корень, степень, скобки) добавляется как новый тип узла,
//     а существующие механизмы LayoutGraph автоматически описывают её связи.
//

namespace AsciiMathParser {
	namespace Core {
		namespace Layout {
			using NodeId = std::uint32_t;

			// Возможные типы пространственных отношений
			enum class RelKind {
				Above,    // «a выше b»
				Below,    // псевдо-запрос (ищется как обратное Above)
				Inside,
				Adjacent,
				LeftOf,
				RightOf
			};

			// Узел графа: текстовый фрагмент или фича (bar, sqrt, bracket, …)
			struct LayoutNode {
				NodeId id;
				Model::Region region;
				std::string role;
			};

			// Ребро графа: направленное отношение между двумя узлами
			struct LayoutEdge {
				NodeId a;
				NodeId b;
				RelKind kind;
			};

			//
			// ░ LayoutGraph
			// ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
			//
			// Универсальная структура для описания топологии выражения.
			// Содержит узлы (всё, что есть на ASCII-плоскости)
			// и направленные рёбра «кто относительно кого где находится».
			//
			class LayoutGraph {
			public:
				LayoutGraph() : nodes{}, edges{} {}

				NodeId AddNode(const Model::Region& region, const std::string& role) {
					const NodeId nid = static_cast<NodeId>(this->nodes.size());
					this->nodes.push_back({ nid, region, role });
					return nid;
				}

				void AddEdge(NodeId a, NodeId b, RelKind kind) {
					this->edges.push_back({ a, b, kind });
				}

				const std::vector<LayoutNode>& Nodes() const noexcept { return this->nodes; }
				const std::vector<LayoutEdge>& Edges() const noexcept { return this->edges; }

				// Быстрый поиск узлов по типу («bar», «text-run», …)
				std::vector<NodeId> FindByRole(const std::string& role) const {
					std::vector<NodeId> res{};
					for (const auto& n : this->nodes) {
						if (n.role == role) res.push_back(n.id);
					}
					return res;
				}

				// Возвращает соседей по заданному типу связи.
				// - Above: все, у кого есть ребро (* → node, Above)
				// - Below: все, у кого есть ребро (node → *, Above)
				std::vector<NodeId> Neighbors(NodeId node, RelKind kind) const {
					std::vector<NodeId> res{};
					if (kind == RelKind::Above) {
						for (const auto& e : this->edges)
							if (e.kind == RelKind::Above && e.b == node) res.push_back(e.a);
						return res;
					}
					if (kind == RelKind::Below) {
						for (const auto& e : this->edges)
							if (e.kind == RelKind::Above && e.a == node) res.push_back(e.b);
						return res;
					}
					// Для симметричных связей (LeftOf и др.)
					for (const auto& e : this->edges)
						if (e.kind == kind && (e.a == node || e.b == node))
							res.push_back(e.a == node ? e.b : e.a);
					return res;
				}

			private:
				std::vector<LayoutNode> nodes;
				std::vector<LayoutEdge> edges;
			};
		}
	}
}
