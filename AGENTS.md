# AGENTS.md

## Project Scope

- The Git root is this repository root.
- Treat paths as repo-relative. Do not use user-specific absolute paths in instructions, examples, or patches.
- Product-specific Power Query source belongs under `Workflows/`.
- Shared Power Query tooling lives beside this repo in `../mcode-pq-tools`.
- Use `pq.project.json` for project-specific paths.

## Skill Routing

- Before any Power Query or M-code work, read and apply [the Power Query M skill](.agents/skills/power-query-m/SKILL.md). This includes reading, creating, debugging, editing, reviewing, documenting, comparing, or validating M code and `.m`/`.pq` source, including workbook-linked source.
- This skill is mandatory even when it is absent from the session's available-skills list. Read its relevant references before working on the source. If the skill is missing or unreadable, report that and stop M-code work rather than silently substituting another skill.
- Keep general Power Query authoring instructions in that skill and its references. Keep ResidentialCare paths, source-of-truth gates, workbook permissions, and project workflows in this file or its linked project documentation.

## Working Rules

- Make only the changes the user requested. Do not bundle related fixes, path changes, refactors, cleanup, or other improvements. If another change seems necessary or useful, explain it and ask the user before editing.
- Do not copy HomeCare business logic into this repository.
- Do not inspect, unzip, parse, diff, compare, or edit Excel workbooks unless explicitly requested.
- Do not create real/full refresh automation until a workbook workflow and source-of-truth process have been approved.
- Keep scaffold, documentation, scripts, and configuration repo-relative.
- For requests to validate, run, monitor, or stop the ResidentialCare refresh sequence, follow `docs/ResidentialCare-Runner-Agent-Instructions.md`.

## Power Query Source Of Truth

- Power Query source files should be committed as `.m` or `.pq` files under `Workflows/`.
- Workbook-linked M-code work must use an approved source-of-truth extraction/sync process before edits are made.
- Do not silently substitute copied, generated, backup, sidecar, or same-stem files as edit targets.
- When the user requests a Power Query review and the exact workbook's extracted `.m`/`.pq` source is missing, stop the review, report the missing source, and ask the user to extract it. Do not inspect the workbook or substitute another unit's, copied, generated, backup, sidecar, or same-stem source.

## Opt-In M-to-Excel Synchronization

- Synchronizing an approved `.m`/`.pq` source file back into an Excel workbook is a deliberate, opt-in operation. Never sync automatically merely because an M file changed.
- Only perform synchronization when the user explicitly instructs it for the named workbook and source file (for example, “sync this M file to Excel”). Treat “edit”, “validate”, or “review” as non-sync instructions.
- Before syncing, identify the exact workbook/source pair, confirm the workbook is closed, validate the M source, and confirm the source-of-truth gate and target path. Do not sync backups, copies, donor workbooks, or same-stem alternatives.
- Use the approved extraction/sync mechanism for the project. The sync operation must replace/update the workbook’s Power Query definition while preserving workbook features, tables, connections, and unrelated queries; do not patch the XLSX ZIP package manually or rebuild the workbook with a general-purpose library.
- After syncing, re-extract the workbook’s Power Query and compare it with the approved M source. Report the exact target, sync status, validation result, and any refresh that was or was not performed.
- Synchronization does not imply refresh. Refresh, save, and close operations require their own explicit approval unless the user’s instruction clearly includes them.

## ResidentialCare Power Query Paths

- When editing an approved workbook-linked M source, review external imports but change their paths only when the user requested path work or approved it separately. Keep the approved client, unit and source-file choices unchanged.
- For approved path work, use one authoritative resolver from `FilePathUrl[FilePath]` through the public-machine `CentriSyncPaths` table. Normalize local/SharePoint paths, handle Excel CELL filename brackets, and use case-insensitive, boundary-safe longest-prefix matching. Derive the unit or organisation root from the resolved workbook folder, then build relative imports from that root.
- For approved path work, remove legacy `Folder` named-table inputs, user-specific absolute paths, fixed-machine roots and competing path resolvers in the edited source. The fixed public-machine CentriSyncPaths bootstrap location is the deliberate exception. Do not fall back to stale loaded path-query outputs.
- Apply the resolver mechanics documented under Flexible ResidentialCare Resolver in `docs/DATExx-Whiddon-Input-Workbook-Path-Setup.md` only to an approved path target. The original three-workbook worksheet-copy procedure and donor approval do not automatically extend to other workbooks. Source edits do not authorize workbook inspection, path-sheet installation, synchronization or refresh; record unverified FilePathUrl prerequisites in the handoff.
- For approved path work, enumerate every external file access in the edited source and verify that each uses a named IMPORT and the authoritative root. Validate the resolver and report any unresolved prerequisites. Do not expand edits to other clients, units or excluded files.

## DATExx Input Workbook Path Setup

- For the approved path-standardisation task covering `1-AllocationExtracted.xlsx`, `2-DemandExtract.xlsx`, and `StaffList Availability.xlsx` under `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/`, follow `docs/DATExx-Whiddon-Input-Workbook-Path-Setup.md`.
- The reference is mandatory for this task and defines the exact workbook scope, donor worksheet, source-of-truth gate, Power Query path standard, and validation process.
- Do not apply it to `CLIENT/DATExx/`, other clients, other units, calculation workbooks, backups, copies, or same-stem files unless the user explicitly expands the scope.

