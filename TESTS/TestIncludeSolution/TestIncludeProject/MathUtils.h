#pragma once

#include <string>

namespace MathUtils
{
    inline std::string Repeat(const std::string& value, const int count)
    {
        std::string result;
        for (int index = 0; index < count; ++index)
        {
            result += value;
        }

        return result;
    }
}
