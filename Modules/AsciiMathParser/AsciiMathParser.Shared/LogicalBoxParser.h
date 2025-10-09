#pragma once
#include <string>
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		struct LineU16 {
			std::u16string text; // UTF-16
			int32_t startInDocument = 0;
		};

		struct LogicalBox {
			int32_t topLine = 0;
			int32_t bottomLine = 0;
			int32_t leftCol = 0;
			int32_t rightCol = 0;
			int32_t topLeftCharIdx = 0;
			int32_t topRightCharIdx = 0;
			int32_t bottomLeftCharIdx = 0;
			int32_t bottomRightCharIdx = 0;
		};

		// Основной API ядра: принимает строки, таб-стопы, отдаёт найденные коробки.
		void FindLogicalBoxes(
			const std::vector<LineU16>& lines,
			int32_t tabSize,
			std::vector<LogicalBox>& outBoxes
		);
	}
}