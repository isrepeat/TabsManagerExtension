#pragma once

namespace AsciiMathParser {
	namespace CLR {
		public value struct LogicalBox {
		public:
			int TopLine;
			int BottomLine;
			int LeftCol;
			int RightCol;
			int TopLeftCharIdx;
			int TopRightCharIdx;
			int BottomLeftCharIdx;
			int BottomRightCharIdx;
		};

		public ref class Bridge abstract sealed {
		public:
			static cli::array<LogicalBox>^ FindLogicalBoxes(
				cli::array<System::String^>^ lines, // строка целиком на каждой линии (UTF-16)
				cli::array<int>^ starts,			// абсолютный Start.Position каждой линии в документе
				int tabSize							// размер таб-стопа
			);
		};
	}
}