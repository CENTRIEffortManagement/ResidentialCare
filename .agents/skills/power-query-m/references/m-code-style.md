# Power Query presentation and commentary

## Steps and staging

- Write Power Query transformations as readable, meaningfully named `let` steps. Give each distinct function or transformation its own step whenever Power Query permits; do not nest a chain of operations when the intermediate results can be named and inspected.
- Present substantial pipelines as separate, meaningfully named staging queries in dependency order: inputs, preparation and validation, historical or analytical staging, allocation or calculation, outputs, then checks. Keep staging queries connection-only in the workbook when supported and workbook changes are authorized.
- Do not hide a multi-stage analysis inside one large record-returning model query merely to expose several outputs. Keep a transformation inside one query only when its steps share one responsibility and one stable row grain.
- Use meaningful query and step names that describe the business transformation. Avoid new auto-generated names such as `Changed Type1`, `Added Custom2`, or `Merged Queries3`; preserve existing names when changing them would break an interface or create unnecessary churn.
- Apply these standards to the requested work. Do not restructure unrelated queries or impose a `...Prepare` stage on every operation.

## Headers and comments

- Every new query, and every materially edited query without a useful header, must have Power Query-compatible `// Query:` and `// Purpose:` lines immediately above its `shared` definition. Add `// Inputs:`, `// Output:`, and `// Notes:` when they clarify dependencies, row grain, required columns, or an approved constraint.
- Comment business rules, non-obvious joins, grain changes, distribution denominators, allocation formulas, unit conversions, validation gates, buffering reasons, and deliberate exceptions. Put the comment immediately before the code it explains.
- Do not narrate obvious mechanical steps line by line. Keep comments factual and current, and remove or update comments that no longer match the code.
- Use only Power Query-compatible `//` comments inside M source. Do not insert Markdown headings, bullets, horizontal rules, or fenced blocks into an M file. Short `// --- Topic ---` separators are allowed when they materially improve navigation in a long query.

## Source review

- Review M syntax, identifier quoting, query references, parentheses, commas, and each final `in` target.
- Review query headers, step names, comments, dependency order, row grain, and visible output/check interfaces before reporting completion.
- Expose required reconciliation, completeness, uniqueness, or allocation validation as a clearly named downstream check query. Output queries must not silently publish results after a required validation failure. This does not authorize adding unrequested checks or auditing source data.
- State which source checks ran and which runtime behaviour remains unverified. Source validation does not imply that a workbook was synchronized or refreshed.

## Memory and performance review

Always perform this review before reporting substantial M-code changes complete. Substantial changes include changes to pipelines, joins, grouping, row expansion, interval/list processing, buffering, source volume, or shared output/check dependencies. Review the affected dependency chain.

- Review buffer placement and width. Select only the columns needed by consumers before buffering, and verify that every referenced field and join, grouping, sorting or duplicate key remains present. Avoid buffering unused diagnostic text, wide intermediate tables or historical records unnecessarily.
- Check repeated evaluation across validation, diagnostics and published outputs. A shared query or buffer is not a durable cache across separate evaluations. Identify where the same expensive result may be calculated or materialised repeatedly.
- Check row growth through joins and expansions, the grain of each stage, and whether applicable worker, facility and date reductions occur before expensive processing. Preserve the approved population, date rules and validation scope.
- Check nested lists/tables carried into expanded rows, repeated full-list scans per row or shift, and growing accumulators that are repeatedly copied or buffered.
- Report the review outcome, any material risks, and the source checks performed. Distinguish structural findings from measured runtime or memory results; state when those measurements have not been made.

This is a source-code review, not a requirement to add data audits, production check queries or test harnesses. Runtime profiling, workbook inspection, synchronization and refresh retain their existing authorization requirements. Propose optimisations outside the requested edit separately.
