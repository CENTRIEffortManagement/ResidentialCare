# DATExx Input Workbook Path Setup

### Approved Scope

- Apply this procedure only to these workbooks:
  - `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/1-AllocationExtracted.xlsx`
  - `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/2-DemandExtract.xlsx`
  - `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/StaffList Availability.xlsx`
- The exact adjacent `*_PowerQuery.m` files for those three workbooks are the approved M-code work files for this migration and are a narrow exception to the normal `Workflows/` source location.
- Do not apply this procedure to `CLIENT/DATExx/`, another client, another unit, calculation workbooks, backups, copies, old files, or same-stem alternatives unless the user explicitly expands the scope.
- Preserve all existing changes. Never replace a target workbook or M file with a clean copy, backup, donor, generated file, or same-stem file.

### Required Workbook Path Table

- Each target workbook must contain the standard worksheet and Excel table that expose the workbook's saved path to Power Query.
- Copy the complete worksheet from `CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Settings Data.xlsx` using Excel's native worksheet-copy operation or another approved method that preserves embedded Power Query and workbook features.
- Before copying, verify in the donor workbook that:
  - the worksheet contains an Excel table named `FilePathUrl`;
  - the table contains one required column named `FilePath` and exactly one data row;
  - the data cell contains the working workbook-path formula, normally based on `CELL("filename")`;
  - the displayed value identifies the donor workbook's saved path or URL; and
  - the table is not a Power Query output that refresh will overwrite.
- If `Settings Data.xlsx` fails any donor check, stop and report the failed condition. Do not silently choose another donor.
- Use `FilePathUrl` as the canonical spelling. Do not rely on the legacy `FilePAthUrl` spelling for new setup.
- If a target already contains a valid `FilePathUrl` table, do not create a duplicate. If a conflicting worksheet, table, or workbook-level name exists, stop and report it.
- Preserve all existing worksheets, names, tables, connections, queries, formulas, formatting, and workbook properties.
- Save each target under its existing name and location. Do not rebuild the sheet approximately, patch the XLSX ZIP package manually, or use a library that may discard Power Query connections.

Power Query reads the workbook table with:

```powerquery
Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content]
```

### Source-of-Truth Gate

- Keep these roles distinct:
  - source workbook: the exact in-scope `.xlsx` file;
  - approved M work file: its exact adjacent `_PowerQuery.m` file;
  - donor workbook: `2. Calculations/Settings Data.xlsx`, used only as the worksheet source; and
  - sync target: the same exact source workbook after M validation.
- Before changing a target, record its workbook and M-file timestamps and Git status, and treat existing modifications as user work.
- Confirm the target and donor workbooks are closed before file-based editing or syncing.
- Extract current Power Query from the exact target through the approved extraction process and confirm it updates the exact adjacent approved M work file.
- If the workbook is newer, timestamps are unclear, extraction produces another path, or the exact work file is missing, stop for a source-of-truth decision.
- Never use an active editor tab, backup, copied workbook, or same-stem file as a substitute for an exact target.

### Flexible ResidentialCare Resolver

- Each workbook that imports external files must use one authoritative flexible ResidentialCare resolver.
- The resolver must:
  1. read `FilePathUrl[FilePath]` from `Excel.CurrentWorkbook()`;
  2. normalize `/` to `\` and remove a trailing slash;
  3. load `%PUBLIC%\Public Scripts\CentriSyncPaths.xlsx` from its fixed public-machine location;
  4. read the named table `CentriSyncPaths`;
  5. treat `SharepointRootUrl` and `SyncedFolderRootPath` as text;
  6. consider `SharepointRootUrl`, `SharepointRootUrl\Shared Documents`, and `SyncedFolderRootPath` as candidate roots;
  7. use a case-insensitive, boundary-safe longest-prefix match;
  8. replace the matched root with `SyncedFolderRootPath`;
  9. remove the current workbook filename to derive its containing folder; and
  10. fail clearly when no mapping matches.
- Expose one authoritative unit-root table, normally `UnitL1PathTABLE`, plus a simple scalar path such as `Unit1Path` for imports.
- Do not retain competing old and new resolver implementations after migration.
- Use relative imports from the resolved unit root, for example:

```powerquery
File.Contents(Unit1Path & "\1. Input\Demand-MasterRoster Manual Read.xlsx")
```

- Do not add user-specific absolute paths.
- This is a generic path-resolution convention only. Do not copy HomeCare business logic, names, or RouteOpt folder assumptions.

### Import Rules

- Required files, tables, sheets, and columns must fail visibly when missing or malformed. Do not mask required failures with broad `try ... otherwise`, empty-table fallbacks, or `MissingField.UseNull`.
- Import named Excel tables by default with `Source{[Item="<TableName>", Kind="Table"]}[Data]`.
- Use a sheet import only when the source genuinely has no named table and the sheet-based contract has been explicitly confirmed.
- When several queries read the same external workbook, open and buffer it once through a shared import query and reference that query downstream.
- Preserve existing query names, outputs, column names, row grain, and business transformations unless the user separately approves a business-logic change.
- `StaffList Availability.xlsx` currently has no extracted shared queries. Adding the path worksheet does not authorize inventing imports or business queries.

### Target-Specific Requirements

- `1-AllocationExtracted.xlsx`:
  - consolidate its path logic to one flexible ResidentialCare resolver;
  - remove obsolete duplicate or delimiter-based resolvers only after confirming no query depends on them;
  - preserve roster-import and allocation business logic; and
  - convert required external paths to the resolved unit path.
- `2-DemandExtract.xlsx`:
  - replace the legacy local-only `LocalUnitFolder` implementation with the flexible resolver;
  - retain relative imports from `Demand-MasterRoster Manual Read.xlsx` and `2. Calculations/Settings Data.xlsx`; and
  - consolidate repeated opens of the same workbook into a shared buffered import where query interfaces remain unchanged.
- `StaffList Availability.xlsx`:
  - add or verify only the path worksheet and `FilePathUrl` table for the current task; and
  - add Power Query only after a concrete external import and interface are separately confirmed.

### Sync and Validation

- After editing each approved M work file:
  1. validate M syntax and `let ... in` structure;
  2. confirm no user-specific absolute paths remain;
  3. confirm there is one authoritative resolver, one canonical `FilePathUrl` lookup, and no unintended duplicate imports;
  4. review the diff and verify business logic and public query interfaces are unchanged;
  5. sync only the approved adjacent M file into its exact workbook using the approved sync process;
  6. open only that target workbook;
  7. confirm `FilePathUrl[FilePath]` identifies the target, not the donor;
  8. refresh the path query before dependent queries;
  9. confirm imports resolve under the expected `DATExx-Whiddon/UNITS/Unit1` root;
  10. save and close only the workbook opened for validation; and
  11. re-extract Power Query and confirm the workbook contains the approved M code.
- Do not run a batch or full-workflow refresh unless the user separately approves it.
- Report worksheet status, table/formula verification, source-of-truth result, M changes, checks, sync status, refresh result, and remaining decisions separately for each target.
- State explicitly in the completion report that `CLIENT/DATExx/` was not changed.
