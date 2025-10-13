#pragma once
#include "Detect/IFeatureDetector.h"
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		namespace Detect {
			//
			// ░ FractionBarDetector
			// ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
			//
			// Этот класс — простейший «детектор признаков».
			// Он сканирует AsciiGrid построчно и ищет горизонтальные
			// последовательности символов '=' длиной ≥ 3, которые окружены
			// непустыми строками сверху и снизу.
			//
			// → Находит «полосы», которые могут быть дробными чертами.
			// → Возвращает FeatureRegion с типом featureKind = "bar".
			//
			class FractionBarDetector final : public IFeatureDetector {
			public:
				explicit FractionBarDetector() = default;

				// Основной метод интерфейса: возвращает все найденные фичи.
				std::vector<FeatureRegion> Detect(const Model::AsciiGrid& grid) const override {
					std::vector<FeatureRegion> out{};

					// Проходим по всем строкам ASCII-сетки
					for (int y = 0; y < grid.Height(); ++y) {						
						int x = 0;

						while (x < grid.Width()) {
							// Нашли первый символ '='
							if (this->IsHorizontalDash(grid.At(y, x))) {
								const int xStart = x;

								// Продлеваем, пока идёт сплошная линия '='
								while ((x + 1) < grid.Width() && this->IsHorizontalDash(grid.At(y, x + 1))) {
									++x;
								}
								const int xEnd = x;

								// Игнорируем короткие штрихи — считаем только длинные
								if ((xEnd - xStart + 1) >= 3) {
									// Проверяем, что над и под линией есть непустые символы
									if (this->IsRealFracBar(grid, y, xStart, xEnd)) {
										FeatureRegion feat{};
										feat.region = Model::Region{
											Model::SpanY{ y, y },
											Model::SpanX{ xStart, xEnd }
										};
										feat.featureKind = "bar";
										feat.bandY = y;
										feat.bandX1 = xStart;
										feat.bandX2 = xEnd;
										out.push_back(feat);
									}
								}
							}
							++x;
						}
					}
					return out;
				}

			private:
				// Проверяет: является ли символ горизонтальной чертой
				bool IsHorizontalDash(char ch) const {
					return (ch == '=');
				}

				// Проверяет, есть ли непустые символы и выше, и ниже линии.
				// Это помогает отличить дробь от простого декора или подчёркивания.
				bool IsRealFracBar(
					const Model::AsciiGrid& grid,
					int y,
					int x1,
					int x2
				) const {
					bool hasAbove = false;
					bool hasBelow = false;

					for (int x = x1; x <= x2; ++x) {
						if (grid.At(y - 1, x) != ' ') {
							hasAbove = true;
						}
						if (grid.At(y + 1, x) != ' ') {
							hasBelow = true;
						}
						if (hasAbove && hasBelow) {
							break;
						}
					}
					return (hasAbove && hasBelow);
				}
			};
		}
	}
}
