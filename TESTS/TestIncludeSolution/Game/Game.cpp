#include "Logger.h"
#include "Config.h"
#include "../Helpers.Shared/SharedUtils.h"
#if 0
#include "Missing.h"
#endif

#include <iostream>

int main()
{
    std::cout << LocalGameLogger::Owner << ": " << Shared::DescribeContext() << ", build " << Shared::BuildNumber << '\n';
    return 0;
}
