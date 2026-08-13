#include "Logger.h"

#include <string_view>

std::string_view EditorLoggerOwner()
{
    return Shared::LoggerOwner();
}
