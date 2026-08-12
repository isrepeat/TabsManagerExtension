#pragma once

#include <string>
#include <utility>

class Greeter
{
public:
    explicit Greeter(std::string name);

    [[nodiscard]] std::string CreateMessage(int repeatCount) const;

private:
    std::string name_;
};
