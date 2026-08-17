# Operation Hub — repository guidance

## Product and working style

- Call the product **Operation Hub** in conversation and documentation. The repository path and executable remain `kintai-checker` / `KintaiChecker.exe` intentionally.
- Prefer staged, reviewable automation. Favour maintainable Python services over extending VBA where practical.
- Preserve existing behaviour unless the requested change explicitly alters it. Explain the design, validate it, then implement it.

## Shared-worktree safety

- This worktree can be used by multiple agent sessions. Before editing, inspect `git status --short`, recent commits, and recent file modification times.
- Treat all pre-existing uncommitted changes as user or other-session work. Do not overwrite, revert, stage, or build them without clear scope.
- Rebuild the executable only after all related source changes are present; a build only contains the source state at the time it is run.

## Operational constraints

- The jinjer API rate limit is shared by the organisation's credentials. Do not run concurrent bulk fetches. After HTTP 429, wait 10–15 minutes before retrying; the missing monthly cache must be fetched again from the start.
- For project background and current implementation status, use the `operation-hub-context` skill and verify its dated notes against the repository and current data.
