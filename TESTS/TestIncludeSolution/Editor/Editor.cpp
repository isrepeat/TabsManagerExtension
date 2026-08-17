#include "../Helpers.Shared/Logger.h"
#include "../Helpers.Shared/SharedUtils.h"

#include <iostream>

int main()
{
    std::cout << Shared::LoggerOwner() << ": " << Shared::DescribeContext() << '\n';
    return 0;
}
