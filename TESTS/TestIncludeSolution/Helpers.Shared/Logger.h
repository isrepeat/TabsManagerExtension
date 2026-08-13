#pragma once

#include "Context.h"

#include <string_view>

namespace Shared
{
    inline std::string_view LoggerOwner()
    {
        return ContextName;
    }
}
