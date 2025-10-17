#pragma once
#include "../Model/INode.h"
#include "../Model/Grid.h"

#include <unordered_map>
#include <vector>
#include <memory>

namespace AsciiMathParser {
	namespace Core {
		namespace Parse {
			struct IRegionFeature;

			// Кандидат фичи (например дробь, корень и т.д.), найденный детектором внутри региона.
			// Каждый объект полностью самодостаточен: он знает свою область (bbox),
			// какие подрегионы нужно разобрать рекурсивно (subregions),
			// какие строки нужно исключить из токенизации (mapRowToSkipRangesStructural),
			// и как собрать финальный узел из разобранных поддеревьев (assembleFn).
			struct Candidate {
				// Прямоугольная область (bounding-box), занимаемая всей фичей.
				// Для дроби это объединение barRegion + numRegion + denRegion.
				Model::Region bbox;

				// Подрегионы, которые нужно рекурсивно разобрать перед сборкой узла.
				// Для дроби это два региона: числитель (num) и знаменатель (den).
				std::vector<Model::Region> subregions;

				// Карта "строка → X-диапазоны", описывающая только структурные зоны оформления,
				// которые нельзя токенизировать как символы.
				// Например, для дроби сюда попадает только линия бара (y строки бара, x1..x2 её ширина).
				std::unordered_map<int, std::vector<Model::SpanX>> mapRowToSkipRangesStructural;

				// Функция, собирающая финальный узел (INode) из разобранных поддеревьев.
				// При вызове RegionWalker передаёт сюда уже готовые наборы потомков для subregions,
				// а функция возвращает конкретный тип узла (например Fraction, Root, Bracket и т.д.).
				std::function<
					std::ex::unique_ptr<Model::INode>(
						std::vector<std::vector<std::ex::unique_ptr<Model::INode>>>&& subtrees
					)
				> assembleFn;
			};

			struct IRegionFeature {
				virtual ~IRegionFeature() {}

				// Найти кандидатов внутри region (каждый кандидат уже самодостаточен).
				virtual std::vector<Candidate> CollectChildren(
					const Model::AsciiGrid& grid,
					const Model::Region& region
				) const = 0;
			};
		}
	}
}