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

## DATExx Input Workbook Path Setup

- For the approved path-standardisation task covering `1-AllocationExtracted.xlsx`, `2-DemandExtract.xlsx`, and `StaffList Availability.xlsx` under `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/`, follow `docs/DATExx-Whiddon-Input-Workbook-Path-Setup.md`.
- The reference is mandatory for this task and defines the exact workbook scope, donor worksheet, source-of-truth gate, Power Query path standard, and validation process.
- Do not apply it to `CLIENT/DATExx/`, other clients, other units, calculation workbooks, backups, copies, or same-stem files unless the user explicitly expands the scope.

