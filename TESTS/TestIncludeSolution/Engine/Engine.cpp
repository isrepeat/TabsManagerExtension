#include "SharedUtils.h"

#include <iostream>

const char* NestedEngineContext();

int main()
{
    std::cout << Shared::DescribeContext() << ": " << NestedEngineContext() << '\n';
    return 0;
}
