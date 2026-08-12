# User command aliases

- When the user says `залей ченжи`, treat it as an explicit request to commit the current repository changes locally.
- Inspect the working tree and create one or more logically separated commits when appropriate, using concise descriptive commit messages.
- Include only relevant source changes; do not commit secrets, temporary files, build outputs, or unrelated user changes.
- Do not push, publish, open a pull request, or otherwise send the commits to a remote. The user will do that themselves.
- Do not amend or rewrite existing commits unless the user explicitly asks.
