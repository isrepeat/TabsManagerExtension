#include "AppConfig.h"
#include "Greeter.h"

#include <iostream>

int main()
{
    const Greeter greeter{AppConfig::applicationName};
    std::cout << greeter.CreateMessage(2) << '\n';
    return 0;
}
