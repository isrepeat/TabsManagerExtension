#pragma once

#ifndef PROJECT_CONTEXT_NAME
#define PROJECT_CONTEXT_NAME "Helpers.Shared"
#endif

namespace Shared
{
    inline constexpr auto ContextName = PROJECT_CONTEXT_NAME;

#if defined(GAME_PROJECT)
    inline constexpr auto HighlightedProject = "GAME_PROJECT";
    inline constexpr int ProjectContextId = 1;
#elif defined(EDITOR_PROJECT)
    inline constexpr auto HighlightedProject = "EDITOR_PROJECT";
    inline constexpr int ProjectContextId = 2;
#elif defined(ENGINE_PROJECT)
    inline constexpr auto HighlightedProject = "ENGINE_PROJECT";
    inline constexpr int ProjectContextId = 3;
#else
    inline constexpr auto HighlightedProject = "NO_PROJECT_CONTEXT";
    inline constexpr int ProjectContextId = 0;
#endif
}
