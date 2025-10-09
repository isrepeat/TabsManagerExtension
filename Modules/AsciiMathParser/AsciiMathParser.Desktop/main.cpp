#include "FractionBars.h"
#include "FractionAst.h"

#include <iostream>
#include <string>
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


std::string ExtractMathBlock(
	const std::string& text
) {
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

int main() {
	//const std::string text =
	//	"//math_start\n"
	//	"                                                       Σ^{n-1}_{j=1} (y_{j} - ŷ_{j})^2\n"
	//	"                    Σ^{n-1}_{j=1} (y_{j} - ŷ_{j})^2 + =================================\n"
	//	"                                                                 1 - y_{j}\n"
	//	"λ̂ = ∏^{k-1}_{i=1}  =====================================================================  · (ŵ^{i-1}_{k})^2\n"
	//	"                                              x_{k}\n"
	//	"//math_end\n";

	//const std::string text =
	//	"//math_start                   \n"
	//	"				                \n"
	//	"          a + b                \n"
	//	"				       x        \n"
	//	"L =  ========================= \n"
	//	"             z                 \n"
	//	"//math_end                     \n";

	//const std::string text =
	//	"//math_start                   \n"
	//	"				     2y         \n"
	//	"          a + b + ======       \n"
	//	"				     x          \n"
	//	"L =  ========================= \n"
	//	"             z                 \n"
	//	"//math_end                     \n";


	const std::string text =
		"//math_start                           \n"
		"				    2y                  \n"
		"          a + b + ====                 \n"
		"				    x               u   \n"
		"L =  ========================= + ===== \n"
		"                    2k             w   \n"
		"             z  +  ====                \n"
		"                    c                  \n"
		"//math_end                             \n";

	const std::string mathBlock = ExtractMathBlock(text);

	// 1) Вырезать блок между маркерами можно отдельно; для demo кормим сразу весь текст.
	AsciiMathParser::Core::AsciiGrid grid{
		mathBlock
	};

	// 2) Найдём все «бары»
	AsciiMathParser::Core::FractionBars finder{
		grid
	};

	const auto bars = finder.FindBars();

	std::cout << "Bars found: " << bars.size() << "\n";
	for (const auto& b : bars) {
		std::cout
			<< "  y=" << b.y
			<< " x=(" << b.x1 << ".." << b.x2 << ")\n";
	}

	// 3) Построим иерархию
	AsciiMathParser::Core::BarTreeBuilder builder{
		grid
	};

	const auto roots = builder.Build(bars);


	// 4) Дамп дерева
	const std::string dump = builder.DumpTree(roots);
	std::cout << "\nTREE:\n" << dump << "\n";

	// 5) Грубый псевдо-LaTeX
	const std::string latex = builder.RenderPseudoLatex(roots);
	std::cout << "Pseudo-LaTeX:\n" << latex << "\n";

	system("pause");
	return 0;
}
