#include "FractionBars.h"
#include <algorithm>

namespace MathAscii {
	//static constexpr char DASH = '=';

	//AsciiGrid::AsciiGrid(
	//	const std::string& text
	//)
	//	: lines{}
	//	, h{ 0 }
	//	, w{ 0 } {
	//		{
	//			std::string current{};
	//			for (char ch : text) {
	//				if (ch == '\t') {
	//					current.push_back(' ');
	//					current.push_back(' ');
	//					current.push_back(' ');
	//					current.push_back(' ');
	//				}
	//				else if (ch == '\r') {
	//					// игнорируем
	//				}
	//				else if (ch == '\n') {
	//					this->lines.push_back(current);
	//					current.clear();
	//				}
	//				else {
	//					current.push_back(ch);
	//				}
	//			}
	//			if (!current.empty()) {
	//				this->lines.push_back(current);
	//			}
	//		}

	//		this->h = static_cast<int>(this->lines.size());
	//		this->w = 0;
	//		for (const auto& line : this->lines) {
	//			this->w = std::max(this->w, static_cast<int>(line.size()));
	//		}
	//		for (auto& line : this->lines) {
	//			if (static_cast<int>(line.size()) < this->w) {
	//				line.resize(static_cast<std::size_t>(this->w), ' ');
	//			}
	//		}
	//}

	//int AsciiGrid::Height() const {
	//	return this->h;
	//}

	//int AsciiGrid::Width() const {
	//	return this->w;
	//}

	//char AsciiGrid::At(
	//	int y,
	//	int x
	//) const {
	//	if (y < 0 || y >= this->h || x < 0 || x >= this->w) {
	//		return ' ';
	//	}
	//	return this->lines[static_cast<std::size_t>(y)][static_cast<std::size_t>(x)];
	//}


	//FractionBars::FractionBars(
	//	const AsciiGrid& gridRef
	//)
	//	: grid{ gridRef } {
	//}

	//bool FractionBars::IsHorizontalDash(
	//	char ch
	//) const {
	//	if (ch == DASH) {
	//		return true;
	//	}
	//	return false;
	//}

	//bool FractionBars::IsColumnWindowNonSpaceOnRow(
	//	int y,
	//	int x1,
	//	int x2
	//) const {
	//	if (y < 0 || y >= this->grid.Height()) {
	//		return false;
	//	}
	//	for (int x = x1; x <= x2; ++x) {
	//		if (this->grid.At(y, x) != ' ') {
	//			return true;
	//		}
	//	}
	//	return false;
	//}

	//bool FractionBars::IsRowWindowAllSpace(
	//	int y,
	//	int x1,
	//	int x2
	//) const {
	//	if (y < 0 || y >= this->grid.Height()) {
	//		return true;
	//	}
	//	for (int x = x1; x <= x2; ++x) {
	//		if (this->grid.At(y, x) != ' ') {
	//			return false;
	//		}
	//	}
	//	return true;
	//}

	//std::vector<FracBar> FractionBars::FindBars() const {
	//	std::vector<FracBar> result{};
	//	{
	//		const int H = this->grid.Height();
	//		const int W = this->grid.Width();

	//		for (int y = 0; y < H; ++y) {
	//			int x = 0;
	//			while (x < W) {
	//				while (x < W && !this->IsHorizontalDash(this->grid.At(y, x))) {
	//					++x;
	//				}
	//				if (x >= W) {
	//					break;
	//				}

	//				const int start = x;
	//				while (x < W && this->IsHorizontalDash(this->grid.At(y, x))) {
	//					++x;
	//				}
	//				const int end = x - 1;
	//				const int len = end - start + 1;

	//				// Упростив: нет риска спутать с минусом — но порог длины оставим для шумоустойчивости
	//				if (len >= 3) {
	//					FracBar bar{ y, start, end };
	//					if (this->HasSupportAboveBelow(bar)) {
	//						result.push_back(bar);
	//					}
	//				}
	//			}
	//		}
	//	}

	//	std::sort(
	//		result.begin(),
	//		result.end(),
	//		[](
	//			const FracBar& a,
	//			const FracBar& b
	//			) {
	//				if ((a.x2 - a.x1) != (b.x2 - b.x1)) {
	//					return (a.x2 - a.x1) > (b.x2 - b.x1);
	//				}
	//				if (a.y != b.y) {
	//					return a.y < b.y;
	//				}
	//				return a.x1 < b.x1;
	//		}
	//	);

	//	return result;
	//}

	//bool FractionBars::HasSupportAboveBelow(
	//	const FracBar& bar
	//) const {
	//	const int yAbove = bar.y - 1;
	//	const int yBelow = bar.y + 1;
	//	const bool above = this->IsColumnWindowNonSpaceOnRow(yAbove, bar.x1, bar.x2);
	//	const bool below = this->IsColumnWindowNonSpaceOnRow(yBelow, bar.x1, bar.x2);
	//	if (above && below) {
	//		return true;
	//	}
	//	return false;
	//}

	//void FractionBars::ComputeNumDenRects(
	//	const FracBar& bar,
	//	Rect& outNumeratorRect,
	//	Rect& outDenominatorRect
	//) const {
	//	const int x1 = bar.x1;
	//	const int x2 = bar.x2;

	//	int yUp = bar.y - 1;
	//	while (yUp >= 0 && !this->IsRowWindowAllSpace(yUp, x1, x2)) {
	//		--yUp;
	//	}
	//	// теперь yUp указывает на "пустую" или вышел за границу; реальный верх — yUp+1
	//	int numTop = yUp + 1;
	//	int numBottom = bar.y - 1;
	//	if (numTop > numBottom) {
	//		numTop = numBottom = bar.y - 1;
	//	}

	//	int yDown = bar.y + 1;
	//	while (yDown < this->grid.Height() && !this->IsRowWindowAllSpace(yDown, x1, x2)) {
	//		++yDown;
	//	}
	//	// yDown указывает на пустую или за границей; реальный низ — yDown-1
	//	int denTop = bar.y + 1;
	//	int denBottom = yDown - 1;
	//	if (denTop > denBottom) {
	//		denTop = denBottom = bar.y + 1;
	//	}

	//	outNumeratorRect = Rect{ numTop, numBottom, x1, x2 };
	//	outDenominatorRect = Rect{ denTop, denBottom, x1, x2 };
	//}

} // namespace MathAscii