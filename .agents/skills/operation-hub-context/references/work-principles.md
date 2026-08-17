# Durable operating context

## User priorities

- Support attendance checking and invoice-related operations with incremental automation.
- Prefer maintainable external Python services over additional VBA when the choice is open.
- Avoid breaking existing functions. Use a cautious design → confirmation → implementation sequence.

## Concurrent work

Several agent sessions can share this repository. Check Git status, recent commits, and recent file modification times before changing source. Do not use the same shared worktree concurrently without either waiting or explicitly separating worktrees. Build the portable executable only after all related changes have landed.

## jinjer API

Bulk jinjer calls share the organisation's effective rate-limit bucket. Do not run two bulk fetches together. If salary-statement fetching returns HTTP 429, wait 10–15 minutes and rerun the affected month from the beginning because its cache is incomplete.

## Naming

The product display name is **Operation Hub** as of 2026-08-14. Keep legacy path, repository, executable, shortcut, and distribution names unchanged unless the user explicitly requests a migration.
