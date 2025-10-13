#pragma once
#include <algorithm>
#include <string>
#include <vector>

//
// ░ AAAA
// ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 
// 

namespace AsciiMathParser {
	namespace Core {
		namespace Model {
			// Горизонтальный диапазон (по оси X).
			// Описывает отрезок: [x1, x2] включительно.
			struct SpanX {
				int x1;
				int x2;

				// Возвращает ширину диапазона (в символах).
				int Width() const {
					return this->x2 - this->x1 + 1;
				}

				// Вычисляет длину пересечения с другим диапазоном.
				int OverlapLength(const SpanX& other) const {
					const int left = std::max(this->x1, other.x1);
					const int right = std::min(this->x2, other.x2);
					if (right < left) {
						return 0;
					}
					return (right - left + 1);
				}
			};


			// Вертикальный диапазон (по оси Y).
			// Описывает диапазон строк: [y1, y2] включительно.
			struct SpanY {
				int y1;
				int y2;

				// Возвращает высоту диапазона (в строках).
				int Height() const {
					return this->y2 - this->y1 + 1;
				}

				// Проверяет, входит ли заданная координата Y в диапазон.
				bool ContainsY(int y) const {
					if (y < this->y1) {
						return false;
					}
					if (y > this->y2) {
						return false;
					}
					return true;
				}
			};


			// Region - двумерный прямоугольник (область на ASCII-сетке).
			struct Region {
				SpanY rows; // Вертикальные границы (строки)
				SpanX cols; // Горизонтальные границы (столбцы)

				// Ширина области (в символах).
				int Width() const {
					return this->cols.Width();
				}

				// Высота области (в строках).
				int Height() const {
					return this->rows.Height();
				}

				// Проверяет, входит ли конкретная координата Y в диапазон строк области.
				bool ContainsY(int y) const {
					return this->rows.ContainsY(y);
				}

				// Возвращает длину пересечения по X с другой областью.
				int OverlapX(const Region& other) const {
					return this->cols.OverlapLength(other.cols);
				}
			};


			// AsciiGrid — обёртка над многострочным ASCII-текстом.
			// Служит базовым источником символов для всех операций.
			// Позволяет обращаться к конкретному символу (y, x)
			// и знать размеры сетки.
			class AsciiGrid {
			public:
				explicit AsciiGrid(const std::string& text) {
					std::string normilizedLine{};
					for (char ch : text) {
						if (ch == '\t') {
							normilizedLine.push_back(' ');
							normilizedLine.push_back(' ');
							normilizedLine.push_back(' ');
							normilizedLine.push_back(' ');
						}
						else if (ch == '\r') {
							// игнорируем
						}
						else if (ch == '\n') {
							this->lines.push_back(normilizedLine);
							normilizedLine.clear();
						}
						else {
							normilizedLine.push_back(ch);
						}
					}

					if (!normilizedLine.empty()) {
						this->lines.push_back(normilizedLine);
					}
					

					this->w = 0;
					for (const auto& line : this->lines) {
						if (static_cast<int>(line.size()) > this->w) {
							this->w = static_cast<int>(line.size());
						}
					}

					// Дополняем все строки пробелами до максимума.
					for (auto& line : this->lines) {
						if (static_cast<int>(line.size()) < this->w) {
							line.resize(static_cast<std::size_t>(this->w), ' ');
						}
					}

					this->h = static_cast<int>(this->lines.size());
				}

				int Width() const {
					return this->w;
				}

				int Height() const {
					return this->h;
				}

				// Безопасно возвращает символ по координатам.
				// Если координата вне диапазона — возвращает пробел.
				char At(int y, int x) const {
					if (y < 0 || y >= this->h || x < 0 || x >= this->w) {
						return ' ';
					}
					return this->lines[static_cast<std::size_t>(y)][static_cast<std::size_t>(x)];
				}

			private:
				std::vector<std::string> lines; // построчное хранение текста
				int w; // ширина (макс. длина строки)
				int h; // высота (число строк)
			};
		}
	}
}