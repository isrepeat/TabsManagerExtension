#include <Helpers/Text.h>

#include "Detect/FractionBarDetector.h"
#include "Parse/FractionFeature.h"
#include "Parse/IRegionFeature.h"
//#include "Parse/RegionParser.h"
#include "Parse/RegionWalker.h"
#include "Model/INode.h"
#include "Model/Grid.h"


#include <iostream>
#include <format>
#include <regex>


/*

									   e
								2y + =====				[bar1]
									   2
					  a + b + ============ * k			[bar2]
									n
				L =  =============================		[bar3]
								 z

1)

			  Num = "e" (raw text)
			/
	  Frac_1
			\
			  Den = "2" (raw text)

							e
			  Num = "2y + ====" (raw text)
			/               2
	  Frac_2
			\
			  Den = "n" (raw text)

									 e
							  2y + =====
									 2
			  Num = "a + b + =========== * k" (raw text)
			/ 		              n
	  Frac_3
			\
			  Den = "z" (raw text)


2)

			  Num = { Symbol{e} }
			/
	  Frac_1
			\
			  Den = { Symbol{2} }


			  Num = { Symbol{2y}, Symbol{+}, Frac{Num: Symbol{e}; Den: Symbol{2}} }
			/
	  Frac_2
			\
			  Den = { Symbol{n} }


			  Num = { Symbol{a}, Symbol{+}, Symbol{b}, Frac{Num: Frac{Num: Symbol{e}; Den: Symbol{2}}; Den: Symbol{n}}, Symbol{*}, Symbol{k} }
			/
	  Frac_3
			\
			  Den = { Symbol{z} }

*/


enum class FormulaType {
	FromulaFrac1 = 0,
	FromulaFrac2,
	FromulaFrac3,
};

const std::map<FormulaType, std::string> formulaTexts = {
	{ FormulaType::FromulaFrac1,
	"//math_start                   \n"
	"                               \n"
	"            a + b              \n"
	"L =  ========================= \n"
	"             z                 \n"
	"//math_end                     \n" },

	{ FormulaType::FromulaFrac2,
	"//math_start                   \n"
	"                    2y         \n"
	"          a + b + ======       \n"
	"                    n          \n"
	"L =  ========================= \n"
	"                z + h          \n"
	"//math_end                     \n" },

	{ FormulaType::FromulaFrac3,
	"//math_start                   \n"
	"                    2y         \n"
	"          a + b + ====== * k   \n"
	"                    n          \n"
	"L =  ========================= \n"
	"                     e         \n"
	"                z + ===        \n"
	"                     2         \n"
	"//math_end                     \n" }
};


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
	using namespace AsciiMathParser::Core;

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

	Model::Region WholeRegion(const Model::AsciiGrid& grid) {
		return Model::Region{
			Model::SpanX{ 0, grid.Width() - 1 },
			Model::SpanY{ 0, grid.Height() - 1 }
		};
	}
}


namespace AsciiMathParser::Core {
	namespace Dump {
		static std::string Bg256(int idx) {
			if (idx < 0) {
				idx = 0;
			}
			if (idx > 255) {
				idx = 255;
			}
			return "\x1b[48;5;" + std::to_string(idx) + "m";
		}

		static constexpr const char* ANSI_RESET = "\x1b[0m";

		void PrintGridWithOuterBarHighlight(
			const AsciiMathParser::Core::Model::AsciiGrid& grid,
			const std::vector<AsciiMathParser::Core::Detect::FractionBar>& bars,
			int barIdx = -1,
			int numBgColor = 236,
			int denBgColor = 240
		) {
			if (bars.empty()) {
				// просто напечатаем грид без подсветки
				for (int y = 0; y < grid.Height(); ++y) {
					for (int x = 0; x < grid.Width(); ++x) {
						std::cout << grid.At(x, y);
					}
					std::cout << '\n';
				}
				return;
			}

			// внешняя — последняя
			const auto& currentBar = barIdx < 0
				? bars.back()
				: bars.at(barIdx);

			const std::string NUM_BG = Bg256(numBgColor);
			const std::string DEN_BG = Bg256(denBgColor);

			for (int y = 0; y < grid.Height(); ++y) {
				for (int x = 0; x < grid.Width(); ++x) {
					const bool isBar =
						(y == currentBar.barRegion.y) &&
						(x >= currentBar.barRegion.cols.x1) &&
						(x <= currentBar.barRegion.cols.x2);

					const bool inNum =
						(x >= currentBar.numRegion.cols.x1) &&
						(x <= currentBar.numRegion.cols.x2) &&
						(y >= currentBar.numRegion.rows.y1) &&
						(y <= currentBar.numRegion.rows.y2);

					const bool inDen =
						(x >= currentBar.denRegion.cols.x1) &&
						(x <= currentBar.denRegion.cols.x2) &&
						(y >= currentBar.denRegion.rows.y1) &&
						(y <= currentBar.denRegion.rows.y2);

					const char ch = grid.At(x, y);

					if (isBar) {
						std::cout << ch;
					}
					else if (inNum) {
						std::cout << NUM_BG << ch << ANSI_RESET;
					}
					else if (inDen) {
						std::cout << DEN_BG << ch << ANSI_RESET;
					}
					else {
						std::cout << ch;
					}
				}
				std::cout << '\n';
			}
		}

		void Indent(int depth) {
			for (int i = 0; i < depth; ++i) {
				std::cout << "  ";
			}
		}

		void DumpNodes(
			const std::vector<std::unique_ptr<Model::INode>>& nodes,
			int depth = 0
		) {
			for (const auto& n : nodes) {
				if (const auto* f = dynamic_cast<const Model::Fraction*>(n.get())) {
					Dump::Indent(depth);
					std::cout << "[Frac]\n";

					Dump::Indent(depth);
					std::cout << "  Num:\n";
					Dump::DumpNodes(f->Num().nodes, depth + 2);

					Dump::Indent(depth);
					std::cout << "  Den:\n";
					Dump::DumpNodes(f->Den().nodes, depth + 2);
				}
				else if (const auto* s = dynamic_cast<const Model::Symbol*>(n.get())) {
					Dump::Indent(depth);
					std::cout << std::format("[Symbol] '{}'\n", s->Content());
				}
				else {
					Dump::Indent(depth);
					std::cout << "[Unknown node]\n";
				}
			}
		}
	}
}



int main() {
	using namespace AsciiMathParser::Core;

	const std::string mathBlock = LocalHelpers::ExtractMathBlock(
		formulaTexts.at(FormulaType::FromulaFrac3)
	);

	std::cout << "==== Raw block ====\n";
	std::cout << mathBlock << "\n\n";

	// Grid
	auto grid = Model::AsciiGrid{ mathBlock };
	auto wholeRegion = LocalHelpers::WholeRegion(grid);

	auto features = std::vector<std::unique_ptr<Parse::IRegionFeature>>{};
	{
		Detect::FractionBarDetector barDetector{ grid };
		std::vector<Detect::FractionBar> fractionBars = barDetector.DetectBars();

		std::cout << std::format("fractionBars count: {}\n", fractionBars.size());

		//std::cout << "\n==== Visual (outer num/den) ====\n";
		//Dump::PrintGridWithOuterBarHighlight(grid, fractionBars, 1, 236, 236);

		features.push_back(std::make_unique<Parse::FractionFeature>(fractionBars));
	}

	Parse::RegionWalker regionWalker{ std::move(features) };

	auto astRootNodes = regionWalker.ParseRegion(
		grid,
		wholeRegion
	);

	std::cout << "\n==== AST ====\n";
	Dump::DumpNodes(astRootNodes, 0);

	std::cout << "\n";
	system("pause");
	return 0;
}