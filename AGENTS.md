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

## Power Query Import And EXTRACT Queries

- When editing an approved workbook-linked M source, review all external imports in that file and apply the standard ResidentialCare path convention as part of the authorized source edit. Do this proactively; do not wait for a separate path-fix request. Keep the approved client, unit and source-file choices unchanged.
- Use one authoritative resolver from `FilePathUrl[FilePath]` through the public-machine `CentriSyncPaths` table. Normalize local/SharePoint paths, handle Excel CELL filename brackets, and use case-insensitive, boundary-safe longest-prefix matching. Derive the unit or organisation root from the resolved workbook folder, then build relative imports from that root.
- Do not retain legacy `Folder` named-table inputs, user-specific absolute paths, fixed-machine roots or competing path resolvers in the edited source. The fixed public-machine CentriSyncPaths bootstrap location is the deliberate exception. Do not fall back to stale loaded path-query outputs.
- Apply the resolver mechanics documented under Flexible ResidentialCare Resolver in `docs/DATExx-Whiddon-Input-Workbook-Path-Setup.md` to an approved target. The original three-workbook worksheet-copy procedure and donor approval do not automatically extend to other workbooks. Source edits do not authorize workbook inspection, path-sheet installation, synchronization or refresh; record unverified FilePathUrl prerequisites in the handoff.
- Before completion, enumerate every external file access in the edited source and verify that each uses a named IMPORT and the authoritative root. Validate the resolver and report any unresolved prerequisites. Do not expand edits to other clients, units or excluded files.

- Always use a separate, meaningfully named `IMPORT` query to read an external source, even when only one extraction is required. Downstream queries must reference that import.
- Use separate `EXTRACT` queries only when more than one distinct extraction is needed from the same imported source.
- Keep `IMPORT` and `EXTRACT` queries free of business transformations. An `IMPORT` establishes and, when useful, buffers the external source or navigation snapshot; an `EXTRACT` only navigates to and returns one source object unchanged.
- Do not filter rows, select or rename columns, set business types, join, calculate, or validate inside an `IMPORT` or `EXTRACT`. Put each such operation in a separate, meaningfully named downstream staging query such as a scope, preparation, or validation query.
- In particular, apply client, unit, facility, role, date, or other business-scope filters after extraction in an explicit scope/preparation query. Preserve the complete configured key value; do not shorten or derive a partial key unless the approved source contract explicitly requires it.
- For a single extraction, select and validate the required value directly in the meaningfully named destination query. For example, `ShiftDuration` can read the `ShiftDuration` table from `IMPORT Settings Data` and validate the hours value directly, without an intermediate `EXTRACT StandardFTEHours` query. Keep this standard duration in hours.
- This rule governs query structure and naming; the workbook Power Query extraction and synchronization requirements above still apply.

## Power Query Presentation And Commentary

- Every new query, and every materially edited query without a useful header, must have Power Query-compatible `// Query:` and `// Purpose:` lines immediately above its `shared` definition. Add `// Inputs:`, `// Output:`, and `// Notes:` when they clarify dependencies, row grain, required columns, or an approved constraint.
- Present substantial pipelines as separate, meaningfully named staging queries in dependency order: inputs, preparation and validation, historical or analytical staging, allocation or calculation, outputs, then checks. Keep staging queries connection-only in the workbook when supported.
- Do not hide a multi-stage analysis inside one large record-returning model query merely to expose several outputs. Keep a transformation inside one query only when its steps share one responsibility and one stable row grain.
- Use meaningful query and step names that describe the business transformation. Avoid new auto-generated names such as `Changed Type1`, `Added Custom2`, or `Merged Queries3`; preserve existing names when changing them would break an interface or create unnecessary churn.
- Comment business rules, non-obvious joins, grain changes, distribution denominators, allocation formulas, unit conversions, validation gates, buffering reasons, and deliberate exceptions. Put the comment immediately before the code it explains.
- Do not narrate obvious mechanical steps line by line. Keep comments factual and current, and remove or update comments that no longer match the code.
- Use only Power Query-compatible `//` comments inside M source. Do not insert Markdown headings, bullets, horizontal rules, or fenced blocks into an M file. Short `// --- Topic ---` separators are allowed when they materially improve navigation in a long query.
- Expose validation as a clearly named check query when an analysis has reconciliation, completeness, uniqueness, or allocation invariants. Output queries must not silently publish results after a required validation failure.
- Before reporting completion, review the query headers, step names, comments, dependency order, final `in` targets, and visible output/check interfaces as part of M-code validation.

## DATExx Input Workbook Path Setup

- For the approved path-standardisation task covering `1-AllocationExtracted.xlsx`, `2-DemandExtract.xlsx`, and `StaffList Availability.xlsx` under `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/`, follow `docs/DATExx-Whiddon-Input-Workbook-Path-Setup.md`.
- The reference is mandatory for this task and defines the exact workbook scope, donor worksheet, source-of-truth gate, Power Query path standard, and validation process.
- Do not apply it to `CLIENT/DATExx/`, other clients, other units, calculation workbooks, backups, copies, or same-stem files unless the user explicitly expands the scope.

