---
afad: "4.0"
version: "0.62.0"
domain: LIMITATIONS
updated: "2026-05-01"
route:
  keywords: [gridgrind, limitations, limits, constraints, cell count, row count, column count, window, sheet name, memory, oom, apache poi, xlsx, excel, max rows, max columns, max cells, max styles, hyperlinks, formula, row height, column width, zoom]
  questions: ["what are the gridgrind limits", "how many rows does gridgrind support", "how many columns does gridgrind support", "what is the maximum window size", "why does gridgrind reject large windows", "what is the cell limit", "what are excel limits", "what are apache poi limits", "does gridgrind support xls", "what is the sheet name limit", "what is the column width limit", "what is the row height limit", "what is the zoom limit"]
---

# Limitations Registry

**Purpose**: Registry of all hard ceilings enforced by GridGrind, Apache POI 5.5.1, and the
Excel `.xlsx` format. Every entry has a stable ID (`LIM-NNN`), but not every entry has the same
propagation model. Product-enforced limits are expected to stay traceable into validation code and
the user-facing contract surfaces that teach them. Upstream reference ceilings remain a canonical
documentation registry even when GridGrind does not yet surface them through one runtime guard,
catalog summary, or help line.

**Two categories of limits:**
- **Product-enforced entries** — categories that start with `GridGrind`, including
  `GridGrind (enforces Excel limit ...)`. These are operational constraints enforced by
  request-model constructors, selector helpers, or explicit request validation. Violations produce
  a structured `INVALID_REQUEST` error with a precise message before or at the relevant product
  boundary. These entries should stay traceable in the enforcing code and, when surfaced to users,
  in help, catalog, and reference text.
- **Reference ceilings** — categories such as `Excel/POI` and `Excel format`. These are upstream
  structural ceilings of the `.xlsx` format or Apache POI 5.5.1. Some are already enforced on
  relevant public request paths; others remain canonical documentation facts rather than a promise
  of one universal GridGrind preflight guard.

**Primary references:**
- Apache POI 5.5.1 SpreadsheetVersion:
  https://poi.apache.org/apidocs/dev/org/apache/poi/ss/SpreadsheetVersion.html
- Apache POI 5.5.1 busy developers quick guide:
  https://poi.apache.org/components/spreadsheet/quick-guide.html
- Apache POI spreadsheet limitations:
  https://poi.apache.org/components/spreadsheet/limitations.html
- Apache POI REL_5_5_1 `XSSFSheet.java`:
  https://github.com/apache/poi/blob/REL_5_5_1/poi-ooxml/src/main/java/org/apache/poi/xssf/usermodel/XSSFSheet.java
- Microsoft Excel specifications and limits (.xlsx):
  https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3

---

## GridGrind Operational Limits

### LIM-001 — Read Window Cell Count

| Field | Value |
|:------|:------|
| **Category** | GridGrind |
| **Limit** | `rowCount * columnCount` must not exceed 250,000 |
| **Error** | `INVALID_REQUEST` |
| **Message** | `rowCount * columnCount must not exceed 250000 but was {n}` |
| **Applies to** | `GET_WINDOW`, `GET_SHEET_SCHEMA` |
| **Code** | `ExcelReadLimits.MAX_WINDOW_CELLS`; `SelectorSupport.requireWindowSize // LIM-001`; `WorkbookReadCommand.requireWindowWithinLimit // LIM-001` |
| **UX** | `--help` Limits section; `GET_WINDOW` and `GET_SHEET_SCHEMA` catalog summaries |

Excel allows worksheets with up to 1,048,576 rows and 16,384 columns. POI can read arbitrarily
large windows; GridGrind must then serialize the result to JSON. A 1,000 x 1,000 window
produces roughly one million cell objects; at typical containerized JVM heap sizes (128-512 MB)
this exhausts memory during serialization. The 250,000-cell cap (e.g., 500 x 500) prevents
`OutOfMemoryError`, empty response files, and unstructured exit-code-1 failures.

This limit is not from Apache POI or Excel. It may be raised in a future release if response
streaming is introduced.

**Recommended pattern for large sheets:** Use `GET_SHEET_SUMMARY` to discover `lastRowIndex` and
`lastColumnIndex`, then tile the populated region with multiple bounded `GET_WINDOW` reads.

---

### LIM-002 — File Format

| Field | Value |
|:------|:------|
| **Category** | GridGrind |
| **Limit** | `source.path` and `persistence.path` must end in `.xlsx` |
| **Error** | `INVALID_REQUEST` |
| **Message** | `path must point to a .xlsx workbook; .xls, .xlsm, and .xlsb are not supported: {path}` |
| **Applies to** | `source` (EXISTING), `persistence` (SAVE_AS) |
| **Code** | `WorkbookPlan.requireXlsxWorkbookPath` |
| **UX** | `--help` Limits section |

GridGrind intentionally narrows Apache POI's broader OOXML workbook support to plain `.xlsx`.
`.xls` (BIFF8/HSSF), macro-enabled `.xlsm`, and `.xlsb` (binary) are all rejected by the product
contract even though Apache POI can preserve and extract VBA-bearing `.xlsm` packages.

---

### LIM-003 — Sheet Name Contract

| Field | Value |
|:------|:------|
| **Category** | GridGrind (enforces Excel limit) |
| **Limit** | Sheet names must be 1 to 31 characters and must not contain `:` `\` `/` `?` `*` `[` `]`, or begin or end with a single quote |
| **Error** | `INVALID_REQUEST` |
| **Message** | `sheetName must not exceed 31 characters: {name}` or `sheetName contains invalid Excel character ':' at position 4: Bad:Name` |
| **Applies to** | All operations with a `sheetName` field |
| **Code** | `ExcelSheetNames.requireValid`; `MutationAction.Validation.requireSheetName`; `ExcelWorkbookSheetSupport.requireSheetName` |
| **UX** | `--help` Limits section; `ENSURE_SHEET` and `RENAME_SHEET` catalog summaries |

The 31-character ceiling and reserved-character rules are Excel hard limits. GridGrind validates
them at request parse time so the error is structured and reported before any workbook work
begins, instead of surfacing a raw Apache POI sheet-creation exception later in execution.

---

### LIM-004 — Column Width

| Field | Value |
|:------|:------|
| **Category** | GridGrind (enforces Excel limit) |
| **Limit** | `widthCharacters` and `sheetDefaults.defaultColumnWidth` must be > 0 and ≤ 255 |
| **Error** | `INVALID_REQUEST` |
| **Messages** | `widthCharacters must not exceed 255.0 (Excel column width limit): got {n}` / `widthCharacters is too small to produce a visible Excel column width: got {n}` |
| **Applies to** | `SET_COLUMN_WIDTH`, `SET_SHEET_PRESENTATION.sheetDefaults.defaultColumnWidth` |
| **Code** | `MutationAction.Validation.requireColumnWidthCharacters`; `ExcelSheetLayoutLimits.requireColumnWidthCharacters`; `ExcelSheetLayoutLimits.requireDefaultColumnWidth` |
| **UX** | `--help` Limits section; `SET_COLUMN_WIDTH` and `SET_SHEET_PRESENTATION` catalog summaries |

255 character units is the Excel column width ceiling. POI converts explicit column widths to
internal units with `round(widthCharacters * 256)`. Values that round to zero are also rejected.
Sheet-wide default column widths are integer character counts and use the same Excel ceiling.
`GET_SHEET_LAYOUT` remains factual on reopen, so malformed positive persisted explicit column
widths can still be reported above `255` instead of being clamped.

---

### LIM-005 — Row Height

| Field | Value |
|:------|:------|
| **Category** | GridGrind (enforces Excel limit) |
| **Limit** | `heightPoints` and `sheetDefaults.defaultRowHeightPoints` must be > 0 and ≤ 409.0 |
| **Error** | `INVALID_REQUEST` |
| **Messages** | `heightPoints must not exceed 409.0 (Excel row height limit): got {n}` / `heightPoints is too small to produce a visible Excel row height: {n}` |
| **Applies to** | `SET_ROW_HEIGHT`, `SET_SHEET_PRESENTATION.sheetDefaults.defaultRowHeightPoints` |
| **Code** | `MutationAction.Validation.requireRowHeightPoints`; `ExcelSheetLayoutLimits.requireRowHeightPoints` |
| **UX** | `--help` Limits section; `SET_ROW_HEIGHT` and `SET_SHEET_PRESENTATION` catalog summaries |

Excel's published maximum row height is 409 points for `.xlsx` worksheets. Apache POI can store a
larger twip value internally, but GridGrind aligns the product contract to the real Excel limit
instead of the wider storage envelope. Values that round to zero twips are also rejected.
`GET_SHEET_LAYOUT` remains factual on reopen, so malformed positive persisted explicit row heights
and default row-height values can still be reported above `409.0` instead of being clamped.

---

### LIM-006 — Duplicate Step IDs

| Field | Value |
|:------|:------|
| **Category** | GridGrind |
| **Limit** | `stepId` values in the `steps` array must be unique |
| **Error** | `INVALID_REQUEST` |
| **Message** | `steps must not contain duplicate stepId values: {id}` |
| **Applies to** | All `steps` entries |
| **Code** | `WorkbookPlan.copySteps // LIM-006` |
| **UX** | Not surfaced in help or catalog (structural protocol rule) |

---

### LIM-007 — GET_CELLS Address List

| Field | Value |
|:------|:------|
| **Category** | GridGrind |
| **Limit** | `addresses` must be non-empty and contain no duplicates |
| **Error** | `INVALID_REQUEST` |
| **Messages** | `addresses must not be empty` / `addresses must not contain duplicates` |
| **Applies to** | `GET_CELLS` |
| **Code** | `SelectorSupport.copyDistinctAddresses // LIM-007`; `ExcelAddressLists.copyNonEmptyDistinctAddresses // LIM-007` |
| **UX** | Not surfaced in help or catalog (structural protocol rule) |

---

## Excel / Apache POI Structural Limits

This section mixes two kinds of upstream-driven constraints:
- workbook-shape ceilings from the `.xlsx` format itself. Many, such as max rows, max columns, max
  text length, max styles, and max function arguments, are reflected in
  `SpreadsheetVersion.EXCEL2007` in Apache POI 5.5.1.
- product-owned safety limits that GridGrind adds because Apache POI 5.5.1 does not reliably
  normalize some structural edits.

Where POI 5.5.1 does not expose a dedicated constant (for example, hyperlink count, formula
length, or nested-function depth), this registry cites the published Excel limit directly instead
of implying a POI-side constant. GridGrind already preflights some workbook-shape ceilings where
the public request model exposes direct row/column addresses, spans, or print bands. The
remaining entries in this section should be read as format ceilings, not a promise that every path
fails early before POI sees an oversized workbook shape.

### LIM-008 — Maximum Rows per Worksheet

| Field | Value |
|:------|:------|
| **Category** | GridGrind (enforces Excel limit for addressed row indices and bands) |
| **Limit** | 1,048,576 rows (2^20) |
| **POI constant** | `SpreadsheetVersion.EXCEL2007.getMaxRows()` returns `0x100000` |
| **Excel spec** | 1,048,576 |
| **Code** | `ExcelRowSpan.MAX_ROW_INDEX`; `SelectorSupport.requireRowIndexWithinBounds // LIM-008`; `MutationAction.Validation.requireRowIndex // LIM-008`; `PrintTitleRowsInput.MAX_ROW_INDEX` |
| **UX** | Structured `INVALID_REQUEST` bounds failures for row selectors, structural edits, and print-title row bands |

Last valid zero-based row index: 1,048,575.

GridGrind already rejects out-of-bounds addressed rows on the public request path for selectors,
row-band operations, row-structural edits, and print-title row bands instead of deferring those
errors to Apache POI.

---

### LIM-009 — Maximum Columns per Worksheet

| Field | Value |
|:------|:------|
| **Category** | GridGrind (enforces Excel limit for addressed column indices and bands) |
| **Limit** | 16,384 columns (2^14, column XFD) |
| **POI constant** | `SpreadsheetVersion.EXCEL2007.getMaxColumns()` returns `0x4000` |
| **Excel spec** | 16,384 |
| **Code** | `ExcelColumnSpan.MAX_COLUMN_INDEX`; `SelectorSupport.requireColumnIndexWithinBounds // LIM-009`; `MutationAction.Validation.requireColumnIndex // LIM-009`; `PrintTitleColumnsInput.MAX_COLUMN_INDEX` |
| **UX** | Structured `INVALID_REQUEST` bounds failures for column selectors, structural edits, and print-title column bands |

Last valid zero-based column index: 16,383.

GridGrind already rejects out-of-bounds addressed columns on the public request path for
selectors, column-band operations, column-structural edits, and print-title column bands.

---

### LIM-010 — Maximum Text Characters per Cell

| Field | Value |
|:------|:------|
| **Category** | Excel/POI |
| **Limit** | 32,767 characters |
| **POI constant** | `SpreadsheetVersion.EXCEL2007.getMaxTextLength()` returns `32767` |
| **Excel spec** | 32,767 |

---

### LIM-011 — Maximum Cell Styles per Workbook

| Field | Value |
|:------|:------|
| **Category** | Excel/POI |
| **Limit** | 64,000 (POI cap); Excel spec: 65,490 |
| **POI constant** | `SpreadsheetVersion.EXCEL2007.getMaxCellStyles()` returns `64000` |
| **Excel spec** | 65,490 |

Note: POI's cap (64,000) is more conservative than Excel's (65,490). The POI figure is the
effective limit when writing via GridGrind.

---

### LIM-012 — Maximum Hyperlinks per Worksheet

| Field | Value |
|:------|:------|
| **Category** | Excel format |
| **Limit** | 65,530 |
| **Excel spec** | 65,530 |

---

### LIM-013 — Maximum Formula Length

| Field | Value |
|:------|:------|
| **Category** | Excel format |
| **Limit** | 8,192 characters |
| **Excel spec** | 8,192 |

---

### LIM-014 — Maximum Nested Function Levels

| Field | Value |
|:------|:------|
| **Category** | Excel format |
| **Limit** | 64 |
| **Excel spec** | 64 |

---

### LIM-015 — Maximum Function Arguments

| Field | Value |
|:------|:------|
| **Category** | Excel/POI |
| **Limit** | 255 |
| **POI constant** | `SpreadsheetVersion.EXCEL2007.getMaxFunctionArgs()` returns `255` |
| **Excel spec** | 255 |

---

### LIM-016 — Structural Edits That Would Move Unsupported Owned Structures

| Field | Value |
|:------|:------|
| **Category** | GridGrind |
| **Limit** | Structural inserts are rejected when they would move a table or sheet-owned autofilter; structural deletes and shifts are also rejected when they would move or truncate a data validation |
| **Error** | `INVALID_REQUEST` |
| **Message** | Product-owned `INSERT_ROWS`, `DELETE_ROWS`, `SHIFT_ROWS`, `INSERT_COLUMNS`, `DELETE_COLUMNS`, or `SHIFT_COLUMNS` message naming the affected structure and sheet |
| **Applies to** | `INSERT_ROWS`, `DELETE_ROWS`, `SHIFT_ROWS`, `INSERT_COLUMNS`, `DELETE_COLUMNS`, `SHIFT_COLUMNS` |
| **Code** | `ExcelRowColumnStructureController.rejectAffectedRowStructuresFor*`; `ExcelRowColumnStructureController.rejectAffectedColumnStructuresFor*` |
| **UX** | `--help` Limits section; structural-edit catalog summaries |

Apache POI 5.5.1 does not reliably normalize table refs or sheet autofilter refs when rows or
columns are structurally inserted, deleted, or shifted, and it still leaves data-validation
coverage unsafe for delete and shift flows. GridGrind now rebuilds data validations across row and
column inserts, but it still rejects the remaining unsafe structural edits instead of persisting
stale workbook XML or silently moving only part of the owned structure.

The rejection is precise and structure-owned. Example messages include:
- `INSERT_ROWS cannot move table 'BudgetTable' on sheet 'Budget'; row structural edits that would move tables are not supported`
- `DELETE_COLUMNS cannot move sheet autofilter A1:C10 on sheet 'Budget'; column structural edits that would move or truncate sheet autofilters are not supported`

---

### LIM-017 — Column Structural Edits While Formulas Are Present

| Field | Value |
|:------|:------|
| **Category** | GridGrind |
| **Limit** | Column structural edits are rejected when the workbook contains any formula cells or formula-defined named ranges |
| **Error** | `INVALID_REQUEST` |
| **Message** | `INSERT_COLUMNS`, `DELETE_COLUMNS`, or `SHIFT_COLUMNS` product-owned message explaining that Apache POI leaves some column references stale |
| **Applies to** | `INSERT_COLUMNS`, `DELETE_COLUMNS`, `SHIFT_COLUMNS` |
| **Code** | `ExcelRowColumnStructureController.rejectFormulaBearingWorkbookForColumnEdit` |
| **UX** | `--help` Limits section; structural-edit catalog summaries |

Apache POI 5.5.1 updates some column references during `shiftColumns(...)`, but not all of them.
In direct probes against formula-bearing workbooks, formulas such as `A2&B2` were rewritten while
others such as `SUM(B2:B4)` remained stale after the same column move. Formula-defined names have
the same risk surface. GridGrind therefore rejects column structural edits whenever formulas or
formula-defined names are present instead of persisting mixed-reference drift.

Example message:
- `SHIFT_COLUMNS cannot run while workbook formulas are present; Apache POI leaves some column references stale during column structural edits`

---

### LIM-018 — Destructive Structural Edits Against Range-Backed Named Ranges

| Field | Value |
|:------|:------|
| **Category** | GridGrind |
| **Limit** | `DELETE_ROWS`, `SHIFT_ROWS`, `DELETE_COLUMNS`, and `SHIFT_COLUMNS` are rejected when they would truncate a range-backed named range or partially move / overwrite one outside the moved band |
| **Error** | `INVALID_REQUEST` |
| **Message** | Product-owned `DELETE_ROWS`, `SHIFT_ROWS`, `DELETE_COLUMNS`, or `SHIFT_COLUMNS` message naming the affected named range and sheet |
| **Applies to** | `DELETE_ROWS`, `SHIFT_ROWS`, `DELETE_COLUMNS`, `SHIFT_COLUMNS` |
| **Code** | `ExcelRowColumnStructureController.rejectDestructiveNamedRangesForRowDelete`; `ExcelRowColumnStructureController.rejectDestructiveNamedRangesForRowShift`; `ExcelRowColumnStructureController.rejectDestructiveNamedRangesForColumnDelete`; `ExcelRowColumnStructureController.rejectDestructiveNamedRangesForColumnShift` |
| **UX** | `--help` Limits section; structural-edit catalog summaries |

Apache POI can safely translate a range-backed named range when the named range is fully contained
inside the moved row or column band, or when it is completely untouched by the edit. Direct probes
against `deleteRows(...)`, `shiftRows(...)`, `deleteColumns(...)`, and `shiftColumns(...)` showed
that destructive overlap produces broken formulas such as `I!#REF!:I!#REF!` and
`I!#REF!:I!$B$1`. GridGrind therefore rejects only the destructive cases instead of persisting
invalid named-range formulas that can later crash workbook rewrites.

Example messages:
- `DELETE_ROWS cannot move named range 'BudgetWindow' on sheet 'Budget'; row structural edits that would truncate range-backed named ranges are not supported`
- `SHIFT_COLUMNS cannot move named range 'BudgetWindow' on sheet 'Budget'; column structural edits that would overwrite or partially move range-backed named ranges are not supported`

---

### LIM-019 — Event-Read Execution Mode Scope

| Field | Value |
|:------|:------|
| **Category** | GridGrind |
| **Limit** | `execution.mode.readMode=EVENT_READ` supports only inspection steps, and only `GET_WORKBOOK_SUMMARY` plus `GET_SHEET_SUMMARY` queries |
| **Error** | `INVALID_REQUEST` |
| **Message** | `execution.mode.readMode=EVENT_READ supports inspection steps only; unsupported step kind: {kind}`, `execution.mode.readMode=EVENT_READ does not support assertion steps`, or `execution.mode.readMode=EVENT_READ supports GET_WORKBOOK_SUMMARY and GET_SHEET_SUMMARY only; unsupported read type: {type}` |
| **Applies to** | top-level `execution.mode.readMode`, `steps[]` |
| **Code** | `DefaultGridGrindRequestExecutor.executionModeFailure`; `ExcelEventWorkbookReader.apply` |
| **UX** | `--help` Limits section; request-shape docs; `execution.mode` docs; `executionModeInputType` catalog summary |

`EVENT_READ` is the low-memory summary reader backed by POI's XSSF event model. It does not
materialize the full workbook object graph, so GridGrind restricts it to workbook and sheet
summary reads only. Unsupported factual reads and analysis reads are rejected at request-validation
time instead of silently falling back to the normal in-memory executor.

---

### LIM-020 — Streaming-Write Execution Mode Scope

| Field | Value |
|:------|:------|
| **Category** | GridGrind |
| **Limit** | `execution.mode.writeMode=STREAMING_WRITE` requires `source.type=NEW`, limits mutation actions to `ENSURE_SHEET` and `APPEND_ROW`, requires `execution.calculation.strategy=DO_NOT_CALCULATE`, allows `markRecalculateOnOpen=true`, requires `ENSURE_SHEET` before any append/assertion/inspection work, and requires at least one `ENSURE_SHEET` mutation |
| **Error** | `INVALID_REQUEST` |
| **Message** | `execution.mode.writeMode=STREAMING_WRITE requires source.type=NEW ...`, `execution.mode.writeMode=STREAMING_WRITE supports ENSURE_SHEET and APPEND_ROW only; unsupported mutation action type: {type}`, `execution.mode.writeMode=STREAMING_WRITE requires execution.calculation.strategy=DO_NOT_CALCULATE ...`, `execution.mode.writeMode=STREAMING_WRITE allows execution.calculation.markRecalculateOnOpen=true only when strategy=DO_NOT_CALCULATE ...`, `execution.mode.writeMode=STREAMING_WRITE requires ENSURE_SHEET before any assertion step ...`, `execution.mode.writeMode=STREAMING_WRITE requires ENSURE_SHEET before any inspection step ...`, or `execution.mode.writeMode=STREAMING_WRITE requires at least one ENSURE_SHEET mutation ...` |
| **Applies to** | top-level `execution.mode.writeMode`, `source`, `steps[]` |
| **Code** | `DefaultGridGrindRequestExecutor.executionModeFailure`; `ExcelStreamingWorkbookWriter.apply` |
| **UX** | `--help` Limits section; request-shape docs; `execution.mode` docs; `executionModeInputType` catalog summary |

`STREAMING_WRITE` is the low-memory append-oriented writer backed by POI `SXSSF`. It authors only
new workbooks and does not expose the full XSSF mutation surface. GridGrind validates the reduced
operation contract up front so callers get deterministic structured errors instead of half-written
streaming workbooks or hidden mode fallback.

---

### LIM-021 — Request JSON Size

| Field | Value |
|:------|:------|
| **Category** | GridGrind |
| **Limit** | request JSON must not exceed `16 MiB` (`16777216` bytes) |
| **Error** | `INVALID_REQUEST` |
| **Message** | `Request JSON exceeds the maximum size of 16 MiB (16777216 bytes); move large authored payloads into UTF8_FILE, FILE, or STANDARD_INPUT sources.` |
| **Applies to** | CLI stdin request ingestion, `--request <path>` file ingestion, `GridGrindJson.readRequest(InputStream)`, `GridGrindJson.readRequest(byte[])` |
| **Code** | `GridGrindContractText.requestDocumentLimitBytes`; `GridGrindJson.requireSupportedRequestLength`; request JSON stream constraints in `GridGrindJson` |
| **UX** | `--help` Limits section; request-shape docs; source-backed input docs |

GridGrind accepts large authored values through source-backed inputs (`UTF8_FILE`, `FILE`, and
`STANDARD_INPUT`), not by letting the outer request document grow without bound. The CLI now
rejects oversized request files before execution starts, and the JSON codec enforces the same
document cap on stdin and byte-array entry points so every transport path fails with one stable
product-owned `INVALID_REQUEST` message.

---

### LIM-022 — Sheet Zoom Percentage

| Field | Value |
|:------|:------|
| **Category** | GridGrind (enforces Excel limit) |
| **Limit** | `zoomPercent` must be between 10 and 400 inclusive |
| **Error** | `INVALID_REQUEST` |
| **Message** | `zoomPercent must be between 10 and 400 inclusive: {n}` |
| **Applies to** | `SET_SHEET_ZOOM` |
| **Code** | `MutationAction.Validation.requireZoomPercent // LIM-022`; `ExcelSheetViewSupport.requireZoomPercent // LIM-022`; `GridGrindLayoutSurfaceReports.SheetLayoutReport` |
| **UX** | `--help` Limits section; `SET_SHEET_ZOOM` catalog summary |

Excel exposes worksheet zoom in the 10% to 400% range. GridGrind validates the authored value up
front and reuses the same bound for factual sheet-layout reports so CLI operators and agents see
one stable limit across mutation, help, catalog, and readback surfaces.

---

## Memory and Performance

Apache POI uses the "usermodel" API, which loads the entire workbook into JVM heap memory.
POI's own guidance: *"As long as you have enough main-memory, you should be able to handle
files up to these limits. For huge files using the default POI classes you will likely need a
very large amount of memory."*
(https://poi.apache.org/components/spreadsheet/limitations.html)

**Containerized deployments.** JVM heap ergonomic defaults are typically ~25% of container
memory. A 512 MB container provides roughly 128 MB of heap. Large workbooks or large read
windows consume proportionally more and can exhaust it — which is why LIM-001 exists.

**Streaming.** POI provides SXSSF for streaming writes and event-model APIs for streaming
reads. GridGrind now exposes both as explicit opt-in execution modes:
- `execution.mode.readMode=EVENT_READ` for low-memory `GET_WORKBOOK_SUMMARY` and
  `GET_SHEET_SUMMARY` requests only (`LIM-019`)
- `execution.mode.writeMode=STREAMING_WRITE` for low-memory append-oriented authoring on `NEW`
  workbooks using `ENSURE_SHEET` and `APPEND_ROW` only, with optional
  `execution.calculation.markRecalculateOnOpen=true`
  (`LIM-020`)
- request JSON documents capped at `16 MiB`, with large authored payloads moved to
  `UTF8_FILE`, `FILE`, or `STANDARD_INPUT` sources
  (`LIM-021`)

All other reads and mutations continue to use the normal full-XSSF in-memory executor.

---

## Unsupported Features

| Feature | Status |
|:--------|:-------|
| Macro-enabled OOXML (`.xlsm`) | Out of scope for the shipped GridGrind contract; LIM-002 rejects it even though Apache POI can preserve and extract VBA from `.xlsm` packages. |
| Charts | Supported for factual reads and authored XDDF plot families `AREA`, `AREA_3D`, `BAR`, `BAR_3D`, `DOUGHNUT`, `LINE`, `LINE_3D`, `PIE`, `PIE_3D`, `RADAR`, `SCATTER`, `SURFACE`, and `SURFACE_3D`, including multi-plot combos built from those families. The remaining limitation is loaded-detail fidelity: unsupported loaded plots still surface as `UNSUPPORTED` and are preserved on unrelated edits instead of being authoritatively mutable. |
| Pivot tables | Limited supported surface: factual reads, health analysis, and authored range-, named-range-, and table-backed pivots. |
| `.xls`, `.xlsm`, `.xlsb` | Not supported. See LIM-002. |
| Streaming read/write | Supported only through `execution.mode`: `EVENT_READ` summary reads (`LIM-019`) and `STREAMING_WRITE` append-oriented `NEW` workbook authoring (`LIM-020`). |
| Request JSON size | Capped at `16 MiB`; large authored values belong in `UTF8_FILE`, `FILE`, or `STANDARD_INPUT` sources (`LIM-021`). |
| OOXML encryption and signing | Supported for `.xlsx` package security on the full-XSSF path. `source.security.password` is required for encrypted sources, `GET_PACKAGE_SECURITY` is unavailable in `EVENT_READ`, and persisting mutations to a signed workbook requires explicit `persistence.security.signature` re-signing. |

Apache POI feature coverage: https://poi.apache.org/components/spreadsheet/

---

## Limit Evolution

When a limit value changes:
1. Update the registry entry here (the ID stays stable).
2. For product-enforced entries, update the constant or literal in the enforcement code
   (referenced in the entry's **Code** row).
3. For product-enforced entries, update all UX strings that surface the value
   (referenced in the entry's **UX** row).
4. For reference ceilings, update the cited upstream source and any impacted explanatory text.
5. Update `CHANGELOG.md` under `[Unreleased]` when the public behavior or published contract
   changes.

Upstream POI releases: https://poi.apache.org/changes.html
