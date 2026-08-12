#include "Greeter.h"

#include "MathUtils.h"

Greeter::Greeter(std::string name)
    : name_(std::move(name))
{
}

std::string Greeter::CreateMessage(const int repeatCount) const
{
    return MathUtils::Repeat("Hello from " + name_ + "! ", repeatCount);
}
