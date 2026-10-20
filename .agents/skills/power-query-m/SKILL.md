---
name: power-query-m
description: Use for any Power Query or M-code work, including reading, creating, debugging, editing, reviewing, documenting, comparing, or validating .m and .pq source; tracing query dependencies and output contracts; and working on workbook-linked M source. Apply this skill even when the task describes queries without mentioning a file extension. This skill does not authorize workbook inspection, synchronization, refresh, or data auditing.
---

# Power Query M

Work on M source as code. Preserve the approved business rules, query interfaces, row grain, and source-of-truth boundaries.

## Read the applicable instructions

- Read the root and applicable nested `AGENTS.md` instructions first. They govern project paths, approved edit targets, workbook permissions, extraction/synchronization gates, path resolvers, and refresh workflows.
- For any query creation, edit, or review, read [query structure and import contracts](references/import-output-contracts.md) and [M-code style](references/m-code-style.md). For other M-code work, read whichever reference governs the task.
- These references hold the general Power Query authoring rules. Do not duplicate them in `AGENTS.md` or import another product's workflow or business logic.

## Land source data unchanged

NEVER transform or check data inside an `IMPORT`, `INPUT`, `EXTRACT`, or `LINK` query. Land the data unchanged, then reference it from meaningfully named downstream queries. Prefer `...Prepare` for preparation work where appropriate; use scope, calculation, or validation queries according to their purpose.

`...Prepare` is a naming and organisation preference, not a mandatory intermediate query for every operation. Ask before implementing it as a broader structural change that the user has not requested.

## Work within the requested scope

1. Identify the exact source file, affected queries, requested behaviour, and applicable source-of-truth gate before editing. Do not substitute a copy, backup, sidecar, or same-stem file without approval.
2. Read only the source and supporting instructions needed to trace the affected dependencies, `let` steps, final `in` targets, row grain, and consumers. Search and use bounded reads for large files.
3. Make the smallest change that addresses the request. Preserve existing output names, columns, types, keys, units, and business rules unless changing them is part of the request.
4. Keep checking proportional to the change. Do not add unrequested data audits, check queries, test harnesses, eligibility changes, or fallback logic. Maintain required validation gates and place any approved data checks downstream of landed inputs.
5. Review syntax, query references, headers, step names, comments, dependency order, and output interfaces using the style reference. Always complete the [memory and performance review](references/m-code-style.md#memory-and-performance-review) for substantial M-code changes before reporting completion. Use available source checks appropriate to the change; do not open or refresh a workbook merely to validate source.
6. Report the exact source changed, resulting behaviour, checks performed, and any runtime limitations. Distinguish source edits from workbook synchronization and refresh; follow the project's approval rules for each operation.

## Workbook boundary

Reading or editing M source does not authorize inspecting, unzipping, parsing, comparing, editing, synchronizing, saving, closing, or refreshing an Excel workbook. Follow the repository's explicit permissions and approved extraction/synchronization mechanism. If the required approved source is missing, follow the repository's missing-source procedure rather than deriving source from a workbook or another product.
