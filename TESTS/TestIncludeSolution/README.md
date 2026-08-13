# TestIncludeSolution

Ручная интеграционная фикстура для проверки контекста C++ Shared Items и графа `IncludeDependencyAnalyzerService`.

## Проекты

- `Game`, `Editor` и `Engine` импортируют один физический `Helpers.Shared.vcxitems`.
- Каждый обычный проект задаёт свой именной макрос (`GAME_PROJECT`, `EDITOR_PROJECT` или `ENGINE_PROJECT`) и `PROJECT_CONTEXT_NAME`.
- В `Context.h` активная `#if`-ветка задаёт `HighlightedProject` и `ProjectContextId`. При открытии заголовка через F12 подсветка IntelliSense должна соответствовать проекту-источнику перехода.

## Основные сценарии

1. Открыть `Game/Game.cpp`, `Editor/Editor.cpp` или `Engine/Engine.cpp`, перейти F12 в `SharedUtils`/`ContextName`, затем открыть `Context.h` и проверить активную ветку `GAME_PROJECT`, `EDITOR_PROJECT` или `ENGINE_PROJECT`.
2. В `Game/Game.cpp` инклюд `Logger.h` должен резолвиться в `Game/LocalIncludes/Logger.h`, а `Config.h` — в Shared Items.
3. В `Editor/Editor.cpp` используется относительный путь к shared-заголовку, а в `Editor/GameLink.cpp` — include directory.
4. В `Engine/Nested/Inner.cpp` shared-заголовок открывается по вложенному относительному пути.
5. Закрывать/выгружать по одному из `Game`, `Editor`, `Engine` и проверять, что представления общих файлов остальных проектов сохраняются.
6. `Missing.h` в `Game.cpp` находится под `#if 0`: проект собирается, но анализатор видит нерезолвимый include.

Все значимые `#include` намеренно расположены в первых десяти строках файлов — это соответствует текущему лимиту анализатора.
