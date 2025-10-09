#include "AsciiMathParser.h"
#include "../AsciiMathParser.Shared/LogicalBoxParser.h"
#include <msclr/marshal.h>
#include <Helpers/String.h>

namespace AsciiMathParser {
	namespace CLR {
		cli::array<LogicalBox>^ Bridge::FindLogicalBoxes(
			cli::array<System::String^>^ lines,
			cli::array<int>^ starts,
			int tabSize
		) {
			if (lines == nullptr || starts == nullptr) {
				throw gcnew System::ArgumentNullException("lines/starts");
			}
			if (lines->Length != starts->Length) {
				throw gcnew System::ArgumentException("lines.Length != starts.Length");
			}
			if (tabSize <= 0) {
				tabSize = 4;
			}

			// 1) Сконвертировать вход в вектор ядра
			std::vector<Core::LineU16> in;
			in.reserve(lines->Length);
			for (int i = 0; i < lines->Length; i++) {
				Core::LineU16 L;
				L.text = H::CLR::Text::ToU16(lines[i]);
				L.startInDocument = starts[i];
				in.push_back(std::move(L));
			}

			// 2) Вызвать ядро
			std::vector<Core::LogicalBox> out;
			Core::FindLogicalBoxes(in, tabSize, out);

			// 3) Вернуть .NET-массив
			auto arr = gcnew cli::array<LogicalBox>(static_cast<int>(out.size()));
			for (int i = 0; i < static_cast<int>(out.size()); i++) {
				LogicalBox v;
				const auto& b = out[static_cast<size_t>(i)];
				v.TopLine = b.topLine;
				v.BottomLine = b.bottomLine;
				v.LeftCol = b.leftCol;
				v.RightCol = b.rightCol;
				v.TopLeftCharIdx = b.topLeftCharIdx;
				v.TopRightCharIdx = b.topRightCharIdx;
				v.BottomLeftCharIdx = b.bottomLeftCharIdx;
				v.BottomRightCharIdx = b.bottomRightCharIdx;
				arr[i] = v;
			}
			return arr;
		}
	}
}