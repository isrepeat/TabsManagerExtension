#pragma once
#include "../Model/Grid.h"
#include <string>
#include <vector>
#include <memory>

namespace AsciiMathParser {
	namespace Core {
		namespace Detect {
			struct FeatureRegion {
				Model::Region region;
				std::string featureKind;
				// опционально: сырые данные (x1,x2,y) — позже пригодятся билдеру
				int bandY = -1;
				int bandX1 = -1;
				int bandX2 = -1;
			};

			class IFeatureDetector {
			public:
				virtual ~IFeatureDetector() = default;

				virtual std::vector<FeatureRegion> Detect(const Model::AsciiGrid& grid) const = 0;
			};
		}
	}
}