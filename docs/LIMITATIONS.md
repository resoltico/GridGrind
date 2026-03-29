---
afad: "3.4"
version: "0.12.0"
domain: LIMITATIONS
updated: "2026-03-29"
route:
  keywords: [gridgrind, limitations, limits, constraints, cell count, row count, column count, window, sheet name, memory, oom, apache poi, xlsx, excel, max rows, max columns, max cells, max styles, hyperlinks, formula, row height, column width]
  questions: ["what are the gridgrind limits", "how many rows does gridgrind support", "how many columns does gridgrind support", "what is the maximum window size", "why does gridgrind reject large windows", "what is the cell limit", "what are excel limits", "what are apache poi limits", "does gridgrind support xls", "what is the sheet name limit", "what is the column width limit", "what is the row height limit"]
---

# Limitations Registry

**Purpose**: Registry of all hard ceilings enforced by GridGrind, Apache POI 5.5.1, and the
Excel `.xlsx` format. Every entry has a stable ID (`LIM-NNN`). IDs are referenced in the
validation code that enforces the limit, in catalog summaries, and in help text, so that any
change to a limit value propagates consistently across all three surfaces.

**Two categories of limits:**
- **GridGrind** — operational constraints enforced by GridGrind compact constructors. Violations
  produce a structured `INVALID_REQUEST` error with a precise message before any workbook work.
  These are the limits most likely to change as the product evolves.
- **Excel/POI** — structural ceilings of the `.xlsx` format reflected in Apache POI
  `SpreadsheetVersion.EXCEL2007`. GridGrind does not currently enforce these at request time;
  violations produce undefined POI behavior or corrupt output.

**Primary references:**
- Apache POI 5.5.1 SpreadsheetVersion:
  https://poi.apache.org/apidocs/dev/org/apache/poi/ss/SpreadsheetVersion.html
- Apache POI spreadsheet limitations:
  https://poi.apache.org/components/spreadsheet/limitations.html
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
| **Code** | `WorkbookReadOperation.MAX_WINDOW_CELLS` (protocol); `WorkbookReadCommand.MAX_WINDOW_CELLS` (engine) |
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
| **Code** | `GridGrindRequest.requireXlsxWorkbookPath` |
| **UX** | `--help` Limits section |

GridGrind uses Apache POI XSSF, which implements only the `.xlsx` (OOXML) format. `.xls`
(BIFF8/HSSF), `.xlsm` (macro-enabled), and `.xlsb` (binary) are not supported.

---

### LIM-003 — Sheet Name Length

| Field | Value |
|:------|:------|
| **Category** | GridGrind (enforces Excel limit) |
| **Limit** | Sheet names must be 1 to 31 characters |
| **Error** | `INVALID_REQUEST` |
| **Message** | `sheetName must not exceed 31 characters: {name}` |
| **Applies to** | All operations with a `sheetName` field |
| **Code** | `WorkbookOperation.Validation.requireSheetName`; `ExcelWorkbook.requireSheetName`; `ExcelNamedRangeTarget` compact constructor |
| **UX** | `--help` Limits section; `ENSURE_SHEET` and `RENAME_SHEET` catalog summaries |

The 31-character ceiling is an Excel hard limit (enforced by Excel and Apache POI). GridGrind
validates at request parse time so the error is structured and reported before any workbook
work begins.

---

### LIM-004 — Column Width

| Field | Value |
|:------|:------|
| **Category** | GridGrind (enforces Excel limit) |
| **Limit** | `widthCharacters` must be > 0 and ≤ 255 |
| **Error** | `INVALID_REQUEST` |
| **Messages** | `widthCharacters must be less than or equal to 255.0: {n}` / `widthCharacters is too small to produce a visible Excel column width: {n}` |
| **Applies to** | `SET_COLUMN_WIDTH` |
| **Code** | `WorkbookOperation.Validation.requireColumnWidthCharacters` |
| **UX** | `--help` Limits section; `SET_COLUMN_WIDTH` catalog summary |

255 character units is the Excel column width ceiling. POI converts the value to internal units
with `round(widthCharacters * 256)`. Values that round to zero are also rejected.

---

### LIM-005 — Row Height

| Field | Value |
|:------|:------|
| **Category** | GridGrind (enforces POI storage limit) |
| **Limit** | `heightPoints` must be > 0 and ≤ 1,638.35 |
| **Error** | `INVALID_REQUEST` |
| **Messages** | `heightPoints is too large for Excel row height storage: {n}` / `heightPoints is too small to produce a visible Excel row height: {n}` |
| **Applies to** | `SET_ROW_HEIGHT` |
| **Code** | `WorkbookOperation.Validation.requireRowHeightPoints` |
| **UX** | `--help` Limits section; `SET_ROW_HEIGHT` catalog summary |

POI stores row heights as a 16-bit signed integer in twips (1/20th of a point).
`Short.MAX_VALUE = 32,767` twips equals approximately 1,638.35 points. Values that round to
zero twips are also rejected.

---

### LIM-006 — Duplicate Read Request IDs

| Field | Value |
|:------|:------|
| **Category** | GridGrind |
| **Limit** | `requestId` values in the `reads` array must be unique |
| **Error** | `INVALID_REQUEST` |
| **Message** | `reads must not contain duplicate requestId values: {id}` |
| **Applies to** | All `reads` entries |
| **Code** | `GridGrindRequest.copyReads` |
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
| **Code** | `WorkbookReadOperation.copyAddresses`; `WorkbookReadCommand.copyAddresses` |
| **UX** | Not surfaced in help or catalog (structural protocol rule) |

---

## Excel / Apache POI Structural Limits

These are hard ceilings of the `.xlsx` format. They are reflected in
`SpreadsheetVersion.EXCEL2007` in Apache POI 5.5.1. GridGrind does not currently enforce them
at request time; exceeding them produces undefined POI behavior or a corrupt output file.

### LIM-008 — Maximum Rows per Worksheet

| Field | Value |
|:------|:------|
| **Category** | Excel/POI |
| **Limit** | 1,048,576 rows (2^20) |
| **POI constant** | `SpreadsheetVersion.EXCEL2007.getMaxRows()` returns `0x100000` |
| **Excel spec** | 1,048,576 |

Last valid zero-based row index: 1,048,575.

---

### LIM-009 — Maximum Columns per Worksheet

| Field | Value |
|:------|:------|
| **Category** | Excel/POI |
| **Limit** | 16,384 columns (2^14, column XFD) |
| **POI constant** | `SpreadsheetVersion.EXCEL2007.getMaxColumns()` returns `0x4000` |
| **Excel spec** | 16,384 |

Last valid zero-based column index: 16,383.

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
| **Category** | Excel/POI |
| **Limit** | 65,530 |
| **Excel spec** | 65,530 |

---

### LIM-013 — Maximum Formula Length

| Field | Value |
|:------|:------|
| **Category** | Excel/POI |
| **Limit** | 8,192 characters |
| **Excel spec** | 8,192 |

---

### LIM-014 — Maximum Nested Function Levels

| Field | Value |
|:------|:------|
| **Category** | Excel/POI |
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
reads. GridGrind does not currently use these; all reads load the full workbook. This may
change in a future release for large-file support.

---

## Unsupported Features

| Feature | Status |
|:--------|:-------|
| Macros (VBA/XLM) | Read: preserved. Write: not creatable. |
| Charts | Not supported (read or write). |
| Pivot tables | Limited (XSSF partial support in POI). |
| `.xls`, `.xlsm`, `.xlsb` | Not supported. See LIM-002. |
| Streaming read/write (SXSSF) | Not used. Full workbook loaded into memory. |

Apache POI feature coverage: https://poi.apache.org/components/spreadsheet/

---

## Limit Evolution

When a limit value changes:
1. Update the registry entry here (the ID stays stable).
2. Update the constant or literal in the enforcement code (referenced in the entry's **Code** row).
3. Update all UX strings that surface the value (referenced in the entry's **UX** row).
4. Update `CHANGELOG.md` under `[Unreleased]`.

Upstream POI releases: https://poi.apache.org/changes.html
