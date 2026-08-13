# User command aliases

- When the user says `залей ченжи`, treat it as an explicit request to commit the current repository changes locally.
- Inspect the working tree and create one or more logically separated commits when appropriate, using concise descriptive commit messages.
- Include only relevant source changes; do not commit secrets, temporary files, build outputs, or unrelated user changes.
- Do not push, publish, open a pull request, or otherwise send the commits to a remote. The user will do that themselves.
- Do not amend or rewrite existing commits unless the user explicitly asks.

# Code comments

- Write new and substantially revised code comments predominantly in Russian unless the surrounding file convention or an external API requires English.
- Keep identifiers, API names, setting keys, protocol values, and exact diagnostic text in their original language.

# C# formatting

- Keep a method call on one line when it remains readable. If at least one argument is moved to another line, put every argument on its own line and put the closing parenthesis on a separate line.
- After a completed multiline method call, add one empty line before the next independent statement. Do not add an empty line inside a single expression or immediately before a closing brace.
