#pragma once
#include "Model/Grid.h"
#include <algorithm>
#include <string>
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		static constexpr char DASH = '=';

		// Горизонтальная «черта» дроби.
		struct FracBar {
			int y;  // Координата строки, где расположена черта
			int x1; // Горизонтальные границы черты
			int x2;

			// Преобразует линию в Region высотой 1 (зона самой черты)
			Model::Region AsBandRegion() const {
				return Model::Region{
					Model::SpanY{ this->y, this->y },
					Model::SpanX{ this->x1, this->x2 }
				};
			}
		};

		////
		//// ░ FractionBars (LEGACY)
		//// ░ Детектор «полос» в стиле старого кода. Оставлен для совместимости.
		//// ░ Новая архитектура использует Detect::FractionBarDetector + LayoutBuilder.
		//class FractionBars {
		//public:
		//	explicit FractionBars(const Model::AsciiGrid& grid)
		//		: grid{ grid } {
		//	}

		//	std::vector<FracBar> FindBars() const {
		//		std::vector<FracBar> bars{};

		//		for (int y = 0; y < this->grid.Height(); ++y) {
		//			int x = 0;
		//			while (x < this->grid.Width()) {
		//				if (this->IsHorizontalDash(this->grid.At(y, x))) {
		//					const int x1 = x;
		//					while ((x + 1) < this->grid.Width()) {
		//						if (!this->IsHorizontalDash(this->grid.At(y, x + 1))) {
		//							break;
		//						}
		//						++x;
		//					}
		//					const int x2 = x;

		//					const int length = x2 - x1 + 1;
		//					if (length >= 3) {
		//						FracBar bar{ y, x1, x2 };
		//						if (this->IsRealFracBar(bar)) {
		//							bars.push_back(bar);
		//						}
		//					}
		//				}
		//				++x;
		//			}
		//		}

		//		return bars;
		//	}

		//	bool IsRealFracBar(const FracBar& bar) const {
		//		bool hasAbove = false;
		//		bool hasBelow = false;

		//		for (int x = bar.x1; x <= bar.x2; ++x) {
		//			if (this->grid.At(bar.y - 1, x) != ' ') {
		//				hasAbove = true;
		//			}
		//			if (this->grid.At(bar.y + 1, x) != ' ') {
		//				hasBelow = true;
		//			}
		//			if (hasAbove && hasBelow) {
		//				break;
		//			}
		//		}

		//		if (hasAbove && hasBelow) {
		//			return true;
		//		}
		//		return false;
		//	}

		//	// Теперь считаем bbox числителя/знаменателя через общие утилиты Geometry.
		//	void ComputeNumDenRects(
		//		const FracBar& bar,
		//		Model::Region& outNumerator,
		//		Model::Region& outDenominator
		//	) const {
		//		const Model::Region band = bar.AsBandRegion();

		//		const auto above = Model::Geometry::TightBBoxAbove(this->grid, band);
		//		if (above.has_value()) {
		//			outNumerator = above.value();
		//		}
		//		else {
		//			outNumerator = Model::Region{
		//				Model::SpanY{ band.rows.y1, band.rows.y1 - 1 },
		//				Model::SpanX{ band.cols.x1, band.cols.x1 - 1 }
		//			};
		//		}

		//		const auto below = Model::Geometry::TightBBoxBelow(this->grid, band);
		//		if (below.has_value()) {
		//			outDenominator = below.value();
		//		}
		//		else {
		//			outDenominator = Model::Region{
		//				Model::SpanY{ band.rows.y2 + 1, band.rows.y2 },
		//				Model::SpanX{ band.cols.x1, band.cols.x1 - 1 }
		//			};
		//		}
		//	}

		//private:
		//	bool IsHorizontalDash(char ch) const {
		//		if (ch == DASH) {
		//			return true;
		//		}
		//		return false;
		//	}

		//private:
		//	const Model::AsciiGrid& grid;
		//};


		//static constexpr char DASH = '=';

		//// Горизонтальная «черта» дроби.
		//struct FracBar {
		//	int y;  // Координата строки, где расположена черта
		//	int x1; // Горизонтальные границы черты
		//	int x2;

		//	// Преобразует линию в Region высотой 1 (зона самой черты)
		//	Model::Region AsBandRegion() const {
		//		return Model::Region{
		//			Model::SpanY{ this->y, this->y },
		//			Model::SpanX{ this->x1, this->x2 }
		//		};
		//	}
		//};





		////
		//// ░ FractionBars
		//// ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 
		//// 
		//// Детектор дробных черт в ASCII-тексте.
		//// Основная роль этого класса — найти все горизонтальные линии,
		//// которые представляют собой дробные черты (‘====’), и для каждой
		//// вычислить области числителя и знаменателя.
		//class FractionBars {
		//public:
		//	explicit FractionBars(const Model::AsciiGrid& grid)
		//		: grid{ grid } {
		//	}

		//	// Ищет все горизонтальные последовательности '='
		//	// и возвращает список найденных линий (FracBar).
		//	std::vector<FracBar> FindBars() const {
		//		std::vector<FracBar> bars{};

		//		for (int y = 0; y < this->grid.Height(); ++y) {
		//			int x = 0;
		//			while (x < this->grid.Width()) {
		//				if (this->IsHorizontalDash(this->grid.At(y, x))) {
		//					int x1 = x;
		//					while (x + 1 < this->grid.Width() && this->IsHorizontalDash(this->grid.At(y, x + 1))) {
		//						++x;
		//					}
		//					int x2 = x;

		//					const int length = x2 - x1 + 1;
		//					if (length >= 3) { // игнорируем короткие штрихи
		//						FracBar bar{ y, x1, x2 };
		//						if (this->IsRealFracBar(bar)) {
		//							bars.push_back(bar);
		//						}
		//					}
		//				}
		//				++x;
		//			}
		//		}

		//		return bars;
		//	}

		//	// Проверяет, есть ли непустые символы и выше, и ниже данной линии.
		//	// Это помогает отличить реальную дробь от простого декора или подчеркивания.
		//	bool IsRealFracBar(const FracBar& bar) const {
		//		bool hasAbove = false;
		//		bool hasBelow = false;

		//		for (int x = bar.x1; x <= bar.x2; ++x) {
		//			if (this->grid.At(bar.y - 1, x) != ' ') {
		//				hasAbove = true;
		//			}
		//			if (this->grid.At(bar.y + 1, x) != ' ') {
		//				hasBelow = true;
		//			}
		//			if (hasAbove && hasBelow) {
		//				break;
		//			}
		//		}

		//		return (hasAbove && hasBelow);
		//	}

		//	// Для заданной линии вычисляет прямоугольники:
		//	//   - outNumerator   — область над линией (числитель)
		//	//   - outDenominator — область под линией (знаменатель)
		//	//
		//	// Алгоритм идёт вверх и вниз, пока не встретит пустую строку.
		//	void ComputeNumDenRects(
		//		const FracBar& bar,
		//		Model::Region& outNumerator,
		//		Model::Region& outDenominator
		//	) const {
		//		const int w = this->grid.Width();
		//		const int h = this->grid.Height();

		//		// вверх — пока не пустая строка
		//		int yUp = bar.y - 1;
		//		while (yUp >= 0) {
		//			if (this->IsRowRangeAllSpace(yUp, bar.x1, bar.x2)) {
		//				break;
		//			}
		//			--yUp;
		//		}

		//		// вниз — пока не пустая строка
		//		int yDown = bar.y + 1;
		//		while (yDown < h) {
		//			if (this->IsRowRangeAllSpace(yDown, bar.x1, bar.x2)) {
		//				break;
		//			}
		//			++yDown;
		//		}

		//		outNumerator = Model::Region{
		//			Model::SpanY{ yUp + 1, bar.y - 1 },
		//			Model::SpanX{ bar.x1, bar.x2 }
		//		};

		//		outDenominator = Model::Region{
		//			Model::SpanY{ bar.y + 1, yDown - 1 },
		//			Model::SpanX{ bar.x1, bar.x2 }
		//		};
		//	}

		//private:
		//	bool IsHorizontalDash(char ch) const {
		//		if (ch == DASH) {
		//			return true;
		//		}
		//		return false;
		//	}

		//	//// Проверяет, есть ли хоть один непустой символ в диапазоне [x1, x2] на строке y.
		//	//bool IsColumnRangeNonSpaceOnRow(int y, int x1, int x2) const {
		//	//	for (int x = x1; x <= x2; ++x) {
		//	//		const char ch = this->grid.At(y, x);
		//	//		if (ch != ' ') {
		//	//			return true;
		//	//		}
		//	//	}
		//	//	return false;
		//	//}

		//	// Проверяет, все ли символы в диапазоне [x1, x2] на строке y — пробелы.
		//	bool IsRowRangeAllSpace(int y, int x1, int x2) const {
		//		for (int x = x1; x <= x2; ++x) {
		//			const char ch = this->grid.At(y, x);
		//			if (ch != ' ') {
		//				return false;
		//			}
		//		}
		//		return true;
		//	}

		//	//// Проверяет, все ли символы в диапазоне [x1, x2] на строке y это '='.
		//	//bool IsRowRangeAllDash(int y, int x1, int x2) const {
		//	//	for (int x = x1; x <= x2; ++x) {
		//	//		const char ch = this->grid.At(y, x);
		//	//		if (!this->IsHorizontalDash(ch)) {
		//	//			return false;
		//	//		}
		//	//	}
		//	//	return true;
		//	//}

		//private:
		//	const Model::AsciiGrid& grid;
		//};
	}
}