# AGENTS.md

## Project Scope

- The Git root is this repository root.
- Treat paths as repo-relative. Do not use user-specific absolute paths in instructions, examples, or patches.
- Product-specific Power Query source belongs under `Workflows/`.
- Shared Power Query tooling lives beside this repo in `../mcode-pq-tools`.
- Use `pq.project.json` for project-specific paths.

## Working Rules

- Do not copy HomeCare business logic into this repository.
- Do not inspect, unzip, parse, diff, compare, or edit Excel workbooks unless explicitly requested.
- Do not create real/full refresh automation until a workbook workflow and source-of-truth process have been approved.
- Keep scaffold, documentation, scripts, and configuration repo-relative.

## Power Query Source Of Truth

- Power Query source files should be committed as `.m` or `.pq` files under `Workflows/`.
- Workbook-linked M-code work must use an approved source-of-truth extraction/sync process before edits are made.
- Do not silently substitute copied, generated, backup, sidecar, or same-stem files as edit targets.

