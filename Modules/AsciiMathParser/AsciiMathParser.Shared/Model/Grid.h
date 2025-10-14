#pragma once
#include <algorithm>
#include <string>
#include <vector>

namespace AsciiMathParser {
	namespace Core {
		namespace Model {
			// Горизонтальный диапазон (по оси X).
			// Описывает отрезок: [x1, x2] включительно.
			struct SpanX {
				const int x1;
				const int x2;

				// Возвращает ширину диапазона (в символах).
				int Width() const {
					return this->x2 - this->x1 + 1;
				}

				// Проверяет, входит ли заданная координата X в диапазон.
				bool ContainsX(int x) const {
					return this->x1 <= x && x <= this->x2;
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
				const int y1;
				const int y2;

				// Возвращает высоту диапазона (в строках).
				int Height() const {
					return this->y2 - this->y1 + 1;
				}

				bool ContainsY(int y) const {
					return this->y1 <= y && y <= this->y2;
				}

				int OverlapLength(const SpanY& other) const {
					const int top = std::max(this->y1, other.y1);
					const int bottom = std::min(this->y2, other.y2);
					return bottom < top ? 0 : (bottom - top + 1);
				}
			};


			// Region - двумерный прямоугольник (область на ASCII-сетке).
			struct Region {
				const SpanX cols; // Горизонтальные границы (столбцы)
				const SpanY rows; // Вертикальные границы (строки)

				// Ширина области (в символах).
				int Width() const {
					return this->cols.Width();
				}

				// Высота области (в строках).
				int Height() const {
					return this->rows.Height();
				}

				// Проверяет, входит ли конкретная координата X в диапазон столцов области.
				bool ContainsX(int x) const {
					return this->cols.ContainsX(x);
				}

				// Проверяет, входит ли конкретная координата Y в диапазон строк области.
				bool ContainsY(int y) const {
					return this->rows.ContainsY(y);
				}

				// Возвращает длину пересечения по X с другой областью.
				int OverlapX(const Region& other) const {
					return this->cols.OverlapLength(other.cols);
				}

				int OverlapY(const Region& other) const {
					return this->rows.OverlapLength(other.rows);
				}
			};


			struct RowRegion {
				const int y;
				const SpanX cols;

				RowRegion(
					const int y,
					const SpanX cols
				)
					: y{ y }
					, cols{ cols } {
				}

				int Width() const {
					return this->cols.Width();
				}

				Region ToRegion() const {
					return Region{
						this->cols,
						SpanY{ this->y, this->y }
					};
				}

				bool ContainsX(const int x) const {
					return this->cols.ContainsX(x);
				}
			};


			struct ColRegion {
				const int x;
				const SpanY rows;

				ColRegion(
					const int x,
					const SpanY rows
				)
					: x{ x }
					, rows{ rows } {
				}

				int Height() const {
					return this->rows.Height();
				}

				bool ContainsY(const int y) const {
					return this->rows.ContainsY(y);
				}

				Region ToRegion() const {
					return Region{
						SpanX{ this->x, this->x },
						this->rows
					};
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
				char At(int x, int y) const {
					if (x < 0 || x >= this->w ||
						y < 0 || y >= this->h
						) {
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