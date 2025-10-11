#include <Helpers/Text.h>

#include "Layout/LayoutBuilder.h"
#include "Layout/LayoutGraph.h"
#include "Detect/FractionBarDetector.h"
#include "Detect/IFeatureDetector.h"
#include "Model/Geometry.h"
#include "Model/Grid.h"

#include <iostream>
#include <format>
#include <regex>


/*
	//math_start
	Some comment...
	
	| 				   2y                   |
	|         a + b + ====                  |
	| 				   x                u   |    |  dy  |
	| L =  ========================= + ==== | +  | ==== |
	| 					2k              w   |    |  dx  |
	| 			 z  +  ====                 |
	| 					c                   |

	//math_end
*/



const std::string text =
"//math_start                   \n"
"				                \n"
"		    a + b               \n"
"L =  ========================= \n"
"             z                 \n"
"//math_end                     \n";

//const std::string text =
//	"//math_start                   \n"
//	"				     2y         \n"
//	"          a + b + ======       \n"
//	"				     x          \n"
//	"L =  ========================= \n"
//	"             z                 \n"
//	"//math_end                     \n";

//const std::string text =
//	"//math_start                           \n"
//	"				    2y                  \n"
//	"          a + b + ====                 \n"
//	"				    x               u   \n"
//	"L =  ========================= + ===== \n"
//	"                    2k             w   \n"
//	"             z  +  ====                \n"
//	"                    c                  \n"
//	"//math_end                             \n";


//const std::string text =
//	"//math_start\n"
//	"                                                       Σ^{n-1}_{j=1} (y_{j} - ŷ_{j})^2\n"
//	"                    Σ^{n-1}_{j=1} (y_{j} - ŷ_{j})^2 + =================================\n"
//	"                                                                 1 - y_{j}\n"
//	"λ̂ = ∏^{k-1}_{i=1}  =====================================================================  · (ŵ^{i-1}_{k})^2\n"
//	"                                              x_{k}\n"
//	"//math_end\n";


namespace LocalHelpers {
	std::string ExtractMathBlock(const std::string& text) {
		// 1. Шаблон: ищем //math_start ... //math_end
		static const std::regex pattern(
			R"(//math_start([\s\S]*?)//math_end)",
			std::regex::ECMAScript
		);

		std::smatch match;
		if (std::regex_search(text, match, pattern)) {
			std::string inner = match[1].str();

			// 2. Убираем ведущие/замыкающие переводы строк и пробелы
			while (!inner.empty() && (inner.front() == '\r' || inner.front() == '\n')) {
				inner.erase(inner.begin());
			}
			while (!inner.empty() && (inner.back() == '\r' || inner.back() == '\n')) {
				inner.pop_back();
			}

			return inner;
		}

		// Если маркеры не найдены — возвращаем оригинал
		return text;
	}


	// Возвращает содержимое узла как строку (без крайних пробелов).
	// Если узел многострочный — строки разделяются символом '|'.
	std::string ExtractNodeContent(
		const AsciiMathParser::Core::Model::AsciiGrid& grid,
		const AsciiMathParser::Core::Layout::LayoutNode& node
	) {
		std::string content{};

		for (int y = node.region.rows.y1; y <= node.region.rows.y2; ++y) {
			for (int x = node.region.cols.x1; x <= node.region.cols.x2; ++x) {
				content.push_back(grid.At(y, x));
			}
			if (y < node.region.rows.y2) {
				content.push_back('|'); // если узел занимает несколько строк
			}
		}

		// Удаляем пробелы по краям
		auto trimFn = [](std::string& s) {
			while (!s.empty() && s.front() == ' ') {
				s.erase(s.begin());
			}
			while (!s.empty() && s.back() == ' ') {
				s.pop_back();
			}
			};

		trimFn(content);
		return content;
	}
}


int main() {
	const std::string mathBlock = LocalHelpers::ExtractMathBlock(text);

	AsciiMathParser::Core::Model::AsciiGrid grid{
		mathBlock
	};

	// -------- Detect --------
	AsciiMathParser::Core::Detect::FractionBarDetector barDetector{};
	const auto features = barDetector.Detect(
		grid
	);

	int barCount = 0;
	for (const auto& f : features) {
		if (f.featureKind == "bar") {
			++barCount;
		}
	}

	std::cout << std::format("Features total: {}\n", features.size());
	std::cout << std::format("  bars: {}\n", barCount);

	for (const auto& f : features) {
		if (f.featureKind == "bar") {
			std::cout << std::format("    bar: y={}, x=({}..{})\n",
				f.bandY,
				f.bandX1,
				f.bandX2
			);
		}
	}

	// -------- Layout --------
	AsciiMathParser::Core::Layout::LayoutBuilderOptions lbOpts{};
	AsciiMathParser::Core::Layout::LayoutBuilder layoutBuilder{
		lbOpts
	};

	auto graph = layoutBuilder.Build(
		grid,
		features
	);

	std::cout << std::format("\nLayoutGraph:\n  nodes: {}\n  edges: {}\n",
		graph.Nodes().size(),
		graph.Edges().size()
	);

	// Выведем все узлы с их ролями и bbox
	for (const auto& n : graph.Nodes()) {
		std::string content = LocalHelpers::ExtractNodeContent(grid, n);

		// Форматированный вывод с фиксированной шириной
		std::cout << std::format(
			"  node#{:02}: role='{:10}'  rows=[{:2}..{:2}]  cols=[{:2}..{:2}]  {:<5} content='{}'\n",
			n.id | H::Text::Color::Gray,
			n.role | H::Text::Color::BrightBlue,
			n.region.rows.y1,
			n.region.rows.y2,
			n.region.cols.x1,
			n.region.cols.x2,
			"", // просто отступ для выравнивания колонок
			(content.empty() ? " " : content) | H::Text::Color::Green
		);
	}


	// Для каждой «bar» покажем соседей Above/Below (что над/под полосой)
	using RelKind = AsciiMathParser::Core::Layout::RelKind;

	const auto barNodes = graph.FindByRole(
		"bar"
	);

	for (const auto nodeId : barNodes) {
		std::cout << std::format("\nbar node#{} relations:\n", nodeId);

		const auto aboveNeighborsIds = graph.Neighbors(
			nodeId,
			RelKind::Above
		);

		const auto belowNeighborsIds = graph.Neighbors(
			nodeId,
			RelKind::Below
		);

		if (!aboveNeighborsIds.empty()) {
			std::cout << "  Above:\n";

			for (const auto aboveNeighborsId : aboveNeighborsIds) {
				const auto& aboveNode = graph.Nodes()[static_cast<std::size_t>(aboveNeighborsId)];
				std::string content = LocalHelpers::ExtractNodeContent(grid, aboveNode);

				std::cout << std::format(
					"  node#{:02}: role='{:10}'  rows=[{:2}..{:2}]  cols=[{:2}..{:2}]  {:<5} content='{}'\n",
					aboveNode.id | H::Text::Color::Gray,
					aboveNode.role | H::Text::Color::BrightBlue,
					aboveNode.region.rows.y1,
					aboveNode.region.rows.y2,
					aboveNode.region.cols.x1,
					aboveNode.region.cols.x2,
					"", // просто отступ для выравнивания колонок
					(content.empty() ? " " : content) | H::Text::Color::Green
				);
			}
		}

		if (!belowNeighborsIds.empty()) {
			std::cout << "  Below:\n";

			for (const auto belowNeighborsId : belowNeighborsIds) {
				const auto& belowNode = graph.Nodes()[static_cast<std::size_t>(belowNeighborsId)];
				std::string content = LocalHelpers::ExtractNodeContent(grid, belowNode);

				std::cout << std::format(
					"  node#{:02}: role='{:10}'  rows=[{:2}..{:2}]  cols=[{:2}..{:2}]  {:<5} content='{}'\n",
					belowNode.id | H::Text::Color::Gray,
					belowNode.role | H::Text::Color::BrightBlue,
					belowNode.region.rows.y1,
					belowNode.region.rows.y2,
					belowNode.region.cols.x1,
					belowNode.region.cols.x2,
					"", // просто отступ для выравнивания колонок
					(content.empty() ? " " : content) | H::Text::Color::Green
				);
			}
		}
	}

	std::cout << "\n";
	system("pause");
	return 0;
}


//int main() {
//	const std::string mathBlock = ExtractMathBlock(text);
//
//	// 1) Вырезать блок между маркерами можно отдельно; для demo кормим сразу весь текст.
//	AsciiMathParser::Core::Model::AsciiGrid grid{
//		mathBlock
//	};
//
//	// 2) Найдём все «бары»
//	AsciiMathParser::Core::FractionBars finder{
//		grid
//	};
//
//	const auto bars = finder.FindBars();
//
//	std::cout << "Bars found: " << bars.size() << "\n";
//	for (const auto& b : bars) {
//		std::cout
//			<< "  y=" << b.y
//			<< " x=(" << b.x1 << ".." << b.x2 << ")\n";
//	}
//
//	//// 3) Построим иерархию
//	//AsciiMathParser::Core::BarTreeBuilder builder{
//	//	grid
//	//};
//
//	//const auto roots = builder.Build(bars);
//
//
//	//// 4) Дамп дерева
//	//const std::string dump = builder.DumpTree(roots);
//	//std::cout << "\nTREE:\n" << dump << "\n";
//
//	//// 5) Грубый псевдо-LaTeX
//	//const std::string latex = builder.RenderPseudoLatex(roots);
//	//std::cout << "Pseudo-LaTeX:\n" << latex << "\n";
//
//	system("pause");
//	return 0;
//}
