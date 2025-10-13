#pragma once
#include <algorithm>
#include <string>
#include <vector>
#include <memory>

namespace AsciiMathParser {
	namespace Core {
		namespace Model {
			struct INode {
				virtual ~INode() = default;
			};


			class Symbol final : public INode {
			public:
				explicit Symbol(std::string content)
					: content{ std::move(content) } {
				}

			public:
				const std::string& GetContent() const {
					return this->content;
				}

			private:
				std::string content;
			};


			class Num {
			public:
				const std::vector<std::unique_ptr<INode>>& GetNodes() const {
					return this->nodes;
				}

				std::vector<std::unique_ptr<INode>>& GetNodes() {
					return this->nodes;
				}

			private:
				std::vector<std::unique_ptr<INode>> nodes;
			};


			class Den {
			public:
				const std::vector<std::unique_ptr<INode>>& GetNodes() const {
					return this->nodes;
				}

				std::vector<std::unique_ptr<INode>>& GetNodes() {
					return this->nodes;
				}

			private:
				std::vector<std::unique_ptr<INode>> nodes;
			};


			class Frac final : public INode {
			public:
				Frac(
					Num num,
					Den den
				)
					: num{ std::move(num) }
					, den{ std::move(den) } {
				}

			public:
				const Num& GetNum() const {
					return this->num;
				}

				const Den& GetDen() const {
					return this->den;
				}

			private:
				Num num;
				Den den;
			};
		}
	}
}