<!--
RETRIEVAL_HINTS:
  keywords: [gridgrind, excel, xlsx, workbook, agent, automation, json, protocol, java, apache-poi, spreadsheet, ai]
  answers: [what is gridgrind, how to use gridgrind, excel automation for ai agents, json workbook api, gridgrind quick start, gridgrind operations, gridgrind install]
  related: [docs/QUICK_REFERENCE.md, docs/OPERATIONS.md, docs/ERRORS.md]
-->

[![GridGrind Art](https://raw.githubusercontent.com/resoltico/GridGrind/main/images/GridGrind.jpg)](https://github.com/resoltico/GridGrind)

-----

# GridGrind

**A fresh grind for every workbook.**

GridGrind is a typed, structured, deterministic `.xlsx` workbook engine with an agent-friendly
JSON protocol. It is designed first for AI-agent and automation workflows, but it is not
agent-only: any caller that can send a JSON request and consume a JSON response can use it.
No low-level library calls, no ad hoc scripting, no brittle string parsing.

---

Alice runs sourcing at High Ground Roasters. Her business lives in Excel: green coffee
inventory by origin, roast batch yields, wholesale price tiers, weekly revenue by SKU. Every
Monday, her AI agent updates those sheets — new arrivals logged, prices recalculated, batch
reports drafted and filed. The agent does not write code to manipulate cells. It sends GridGrind
a JSON request. High Ground's numbers are always current, always correct, always where they need
to be.

Today GridGrind ships as a CLI/container transport around that JSON protocol. It does not yet
ship an MCP server. In other words: GridGrind is already agent-first, but not yet MCP-native.

---

## Running GridGrind

### Container (recommended)

No install required beyond Docker. Works on macOS, Linux, and Windows across both x64 and ARM:

```bash
docker pull ghcr.io/resoltico/gridgrind:latest
```

To pin to a specific release instead of tracking `latest`:

```bash
docker pull ghcr.io/resoltico/gridgrind:0.31.0
```

The container registry retains the last 5 releases. For older versions, download the fat JAR
directly from the [Releases page](https://github.com/resoltico/GridGrind/releases).

Pipe a JSON request to stdin, receive a JSON response on stdout:

```bash
echo '{"source":{"type":"NEW"},"operations":[],"reads":[]}' \
  | docker run -i ghcr.io/resoltico/gridgrind:latest
```

To check the version:

```bash
docker run --rm ghcr.io/resoltico/gridgrind:latest --version
```

To ask the artifact for the current contract directly:

```bash
docker run --rm ghcr.io/resoltico/gridgrind:latest --help
docker run --rm ghcr.io/resoltico/gridgrind:latest --print-request-template
docker run --rm ghcr.io/resoltico/gridgrind:latest --print-protocol-catalog
```

The protocol catalog is machine-readable JSON. Each entry now lists its fields, whether each field
is required or optional, and the exact nested/plain type group accepted by polymorphic fields such
as `value`, `target`, `selection`, `style`, and `scope`. The built-in catalog and request template
are emitted deterministically, so discovery output is stable across repeated runs.

To read a request file and write response and `.xlsx` files back to the host,
mount a working directory:

```bash
docker run -i \
  -v "$(pwd)":/workdir \
  -w /workdir \
  ghcr.io/resoltico/gridgrind:latest \
  --request request.json \
  --response response.json
```

Paths passed in `--request`, `--response`, `source.path`, and `persistence.path` are resolved in
the current execution environment. In the Docker example above, that means container-visible paths
under `/workdir`, so generated `.xlsx` files land back in `$(pwd)` on the host.

### Fat JAR (requires Java 26)

Download the self-contained JAR from the
[Releases page](https://github.com/resoltico/GridGrind/releases/latest) and run it the same
way as the container — stdin/stdout or explicit file paths:

`java -jar gridgrind.jar` uses the ambient shell `java`, not Gradle toolchains. For local
development and release work, make sure `command -v java` resolves to Java 26 and not to the
macOS `/usr/bin/java` stub; see [docs/DEVELOPER_JAVA.md](./docs/DEVELOPER_JAVA.md).

```bash
echo '{"source":{"type":"NEW"},"operations":[],"reads":[]}' \
  | java -jar gridgrind.jar

java -jar gridgrind.jar --request request.json --response response.json
java -jar gridgrind.jar --version
java -jar gridgrind.jar --print-request-template
java -jar gridgrind.jar --print-protocol-catalog
```

### File Workflow

GridGrind has one path model across the CLI, Docker, and the JSON protocol:

- No `--request`: read the JSON request from stdin.
- `--request <path>`: read the JSON request from that file.
- No `--response`: write the JSON response to stdout.
- `--response <path>`: write the JSON response to that file; parent directories are created.
- `source.path`: open an existing workbook from that path.
- `persistence: { "type": "SAVE_AS", "path": ... }`: write a new workbook to that path; parent directories are created.
- `persistence: { "type": "OVERWRITE" }`: write back to `source.path`; `OVERWRITE` does not accept its own path field.

Relative paths in `--request`, `--response`, `source.path`, and `persistence.path` always resolve
from the current working directory. In Docker, that means the container working directory, so the
most predictable workflow is to mount a host directory and set `-w` to that mount point.

---

## How It Works

GridGrind currently supports `.xlsx` workbooks only. Malformed JSON is rejected as
`INVALID_JSON`. Syntactically valid JSON that does not match the protocol model is rejected as
`INVALID_REQUEST_SHAPE` with product-owned messages such as unknown field, unknown type, or wrong
value shape. Parsed requests with invalid business data, such as `.xls`, `.xlsm`, `.xlsb`, or
other non-`.xlsx` workbook paths, are rejected as `INVALID_REQUEST`.

Every request follows the same pipeline:

1. **Operations** run in order — create sheets, write cells, apply styles, evaluate formulas.
2. **Reads** run after all operations succeed — explicit post-mutation introspection and analysis
   operations return only the workbook facts the caller requested.
3. **Persistence** happens last — the workbook is written only after reads succeed.

If any step fails, GridGrind returns a structured error and no file is written. That gives
autonomous callers deterministic failure semantics instead of "error after side effect."

Internally, GridGrind is enforced as `cli -> protocol -> engine`. The public JSON contract lives
in protocol-owned packages grouped by responsibility: `dto`, `operation`, `read`, `catalog`,
`json`, and `exec`. `DefaultGridGrindRequestExecutor` is the single protocol-to-engine
translation boundary into the Apache-POI-backed workbook engine.

---

## Operations

| Operation | What It Does |
|:----------|:-------------|
| `ENSURE_SHEET` | Create a sheet if it does not already exist |
| `RENAME_SHEET` | Rename an existing sheet to a new name |
| `DELETE_SHEET` | Remove an existing sheet (the last remaining sheet or last visible sheet cannot be deleted) |
| `MOVE_SHEET` | Move an existing sheet to a zero-based position |
| `COPY_SHEET` | Copy a sheet into a new visible, unselected sheet at a requested position |
| `SET_ACTIVE_SHEET` | Set the active sheet and ensure it is selected |
| `SET_SELECTED_SHEETS` | Set the selected visible sheet set |
| `SET_SHEET_VISIBILITY` | Set one sheet visibility state |
| `SET_SHEET_PROTECTION` | Enable sheet protection with the supported lock flags |
| `CLEAR_SHEET_PROTECTION` | Disable sheet protection entirely |
| `MERGE_CELLS` | Merge a rectangular A1-style range |
| `UNMERGE_CELLS` | Remove a merged region by exact range match |
| `SET_COLUMN_WIDTH` | Set one or more column widths using character units |
| `SET_ROW_HEIGHT` | Set one or more row heights using point units |
| `INSERT_ROWS` | Insert one or more blank rows before a zero-based row index |
| `DELETE_ROWS` | Delete one inclusive zero-based row band |
| `SHIFT_ROWS` | Move one inclusive zero-based row band by a signed delta |
| `INSERT_COLUMNS` | Insert one or more blank columns before a zero-based column index |
| `DELETE_COLUMNS` | Delete one inclusive zero-based column band |
| `SHIFT_COLUMNS` | Move one inclusive zero-based column band by a signed delta |
| `SET_ROW_VISIBILITY` | Hide or unhide one inclusive zero-based row band |
| `SET_COLUMN_VISIBILITY` | Hide or unhide one inclusive zero-based column band |
| `GROUP_ROWS` | Apply one outline group to one inclusive zero-based row band |
| `UNGROUP_ROWS` | Remove one outline group from one inclusive zero-based row band |
| `GROUP_COLUMNS` | Apply one outline group to one inclusive zero-based column band |
| `UNGROUP_COLUMNS` | Remove one outline group from one inclusive zero-based column band |
| `SET_SHEET_PANE` | Apply an explicit pane state using `NONE`, `FROZEN`, or `SPLIT` |
| `SET_SHEET_ZOOM` | Set the sheet zoom percentage |
| `SET_PRINT_LAYOUT` | Apply the supported print-layout state to one sheet |
| `CLEAR_PRINT_LAYOUT` | Clear the supported print-layout state from one sheet |
| `SET_CELL` | Write a typed value to a single cell without clearing existing style, hyperlink, or comment state |
| `SET_RANGE` | Write a rectangular grid of typed values without clearing existing style, hyperlink, or comment state |
| `CLEAR_RANGE` | Remove values, styles, hyperlinks, and comments from a rectangular range |
| `SET_HYPERLINK` | Attach a hyperlink to a single cell |
| `CLEAR_HYPERLINK` | Remove a hyperlink from a cell; no-op when the cell does not exist |
| `SET_COMMENT` | Attach a plain-text comment to a single cell |
| `CLEAR_COMMENT` | Remove a comment from a cell; no-op when the cell does not exist |
| `APPLY_STYLE` | Apply nested `numberFormat`/`alignment`/`font`/`fill`/`border`/`protection` style patches to a range |
| `SET_DATA_VALIDATION` | Create or replace one data-validation rule over a sheet range |
| `CLEAR_DATA_VALIDATIONS` | Remove data-validation coverage from selected ranges or an entire sheet |
| `SET_CONDITIONAL_FORMATTING` | Create or replace one logical conditional-formatting block over one or more ranges |
| `CLEAR_CONDITIONAL_FORMATTING` | Remove conditional-formatting blocks that intersect selected ranges or an entire sheet |
| `SET_AUTOFILTER` | Create or replace one sheet-level autofilter range |
| `CLEAR_AUTOFILTER` | Remove the sheet-level autofilter range from one sheet |
| `SET_TABLE` | Create or replace one workbook-global table definition |
| `DELETE_TABLE` | Remove one existing workbook-global table by name and expected sheet |
| `SET_NAMED_RANGE` | Create or replace one workbook- or sheet-scoped named range |
| `DELETE_NAMED_RANGE` | Remove one existing workbook- or sheet-scoped named range |
| `APPEND_ROW` | Append a row of typed values after the last value-bearing row |
| `AUTO_SIZE_COLUMNS` | Deterministically size columns to fit their contents |
| `EVALUATE_FORMULAS` | Force formula recalculation before save |
| `FORCE_FORMULA_RECALCULATION_ON_OPEN` | Mark the workbook to recalculate when opened in Excel |

Cell values accept types: `TEXT`, `RICH_TEXT`, `NUMBER`, `BOOLEAN`, `FORMULA`, `DATE`,
`DATE_TIME`, `BLANK`.

Sheet-management operations use strict semantics: missing sheets fail, `MOVE_SHEET.targetIndex`
is zero-based, and `RENAME_SHEET` requires a valid destination name that does not conflict with
another sheet. `COPY_SHEET` creates a new visible, unselected sheet and uses a GridGrind-owned
copy path instead of POI's raw `cloneSheet()` seam, so supported sheet-local content stays
converged while unsupported copy cases such as tables and sheet-scoped formula-defined named
ranges fail explicitly. `SET_ACTIVE_SHEET` and `SET_SELECTED_SHEETS` keep the active tab inside
the selected visible sheet set. `DELETE_SHEET` rejects attempts to delete the last visible sheet,
and `SET_SHEET_VISIBILITY` rejects attempts to hide it. `SET_SHEET_PROTECTION` and
`CLEAR_SHEET_PROTECTION` author and report the exact supported lock-flag set without password
semantics.

Structural layout operations are also strict: `MERGE_CELLS` uses A1-style ranges, `UNMERGE_CELLS`
requires an exact merged-region match, `SET_COLUMN_WIDTH.widthCharacters` is converted to POI
width units with `round(widthCharacters * 256)`, `SET_ROW_HEIGHT.heightPoints` is passed through
as Excel point units, and `SET_SHEET_PANE` uses one explicit pane family. `FROZEN` panes use
split counts plus visible-origin coordinates, while `SPLIT` panes use Excel split offsets in
`1/20` of a point plus visible origin and active-pane state. `SET_SHEET_ZOOM` accepts integers
from `10` through `400`. `SET_PRINT_LAYOUT` covers print area, portrait or landscape orientation,
fit scaling, repeating rows, repeating columns, and plain left/center/right header or footer
text, while `CLEAR_PRINT_LAYOUT` restores the supported default state.

Row and column structural editing is now supported directly. `INSERT_ROWS`, `DELETE_ROWS`,
`SHIFT_ROWS`, `INSERT_COLUMNS`, `DELETE_COLUMNS`, and `SHIFT_COLUMNS` change sheet geometry.
`SET_ROW_VISIBILITY`, `SET_COLUMN_VISIBILITY`, `GROUP_ROWS`, `UNGROUP_ROWS`, `GROUP_COLUMNS`,
and `UNGROUP_COLUMNS` control outline and hidden-state presentation. To prevent silent workbook
drift, row structural edits reject any request that would move or truncate a table, sheet-owned
autofilter, or data validation. Row deletes and shifts also reject destructive interactions with
range-backed named ranges, while still allowing names that are fully contained inside the moved
band or completely untouched by the edit. Column structural edits enforce the same ownership rule,
delete and shift operations reject destructive range-backed named-range rewrites, and all column
structural edits still reject workbooks that contain formulas or formula-defined names because
Apache POI leaves some column references stale after column moves.

`APPLY_STYLE` uses one nested style contract with `numberFormat`, `alignment`, `font`, `fill`,
`border`, and `protection` groups. Alignment patches cover wrap, horizontal or vertical
alignment, text rotation, and indentation. Font patches cover bold, italic, `fontName`, typed
`fontHeight`, `fontColor`, `underline`, and `strikeout`. Fill patches cover solid and patterned
fills with explicit foreground or background colors. Border patches support per-side styles plus
RGB colors, but a border color is only valid when that side has a visible style directly or via
`border.all`, and protection patches cover `locked` and `hiddenFormula`. Style reads in
`GET_CELLS`, `GET_WINDOW`, and `GET_SHEET_SCHEMA` mirror that nested shape; the read-side
`style.font.fontHeight` object reports both exact `twips` and human-friendly `points`.

`RICH_TEXT` is a typed cell-value family rather than a markup convention. It writes an ordered,
non-empty `runs` list, each run carrying plain `text` plus an optional nested font patch
(`fontName`, `fontHeight`, `bold`, `italic`, `underline`, `strikeout`, `fontColor`). Zero-length
runs are rejected. String reads through `GET_CELLS`, `GET_WINDOW`, and `GET_SHEET_SCHEMA` now
return the plain `stringValue` plus optional structured `richText` runs when the stored cell
contains authored rich text.

Authoring metadata is also supported directly. `SET_HYPERLINK` and `CLEAR_HYPERLINK`
work on one cell at a time with typed `URL`, `EMAIL`, `FILE`, or `DOCUMENT` targets. `FILE`
targets now use the field name `path`, accept either plain file paths or `file:` URIs on write,
and are returned as normalized plain path strings on read.
`SET_COMMENT` and `CLEAR_COMMENT` work on one cell at a time with plain-text comment content,
author, and visible state. `SET_NAMED_RANGE` and `DELETE_NAMED_RANGE` work on workbook or sheet
scope with explicit sheet-qualified cell or range targets. Analysis can now return cell
`hyperlink()` and `comment()` metadata plus workbook-level `namedRanges`.

GridGrind now also supports Excel data validation directly. `SET_DATA_VALIDATION` creates or
replaces one rule over an A1-style range, normalizing any overlapping existing rules so the new
rule is authoritative on its target cells. `CLEAR_DATA_VALIDATIONS` removes either the whole
sheet's validation state or just the parts intersecting a selected set of ranges. Supported rule
families include explicit lists, formula-driven lists, whole-number and decimal comparisons,
date/time comparisons, text-length comparisons, and custom formulas. `GET_DATA_VALIDATIONS`
returns either fully modeled definitions or typed `UNSUPPORTED` entries with `kind` and `detail`
metadata when a workbook rule can be detected but not represented as a supported definition.

Conditional formatting is now first-class within the same deterministic request model.
`SET_CONDITIONAL_FORMATTING` writes one logical block with one or more target ranges and an
ordered rule list. Phase A write support covers `FORMULA_RULE` and `CELL_VALUE_RULE`, each with a
typed differential-style payload. `CLEAR_CONDITIONAL_FORMATTING` removes every stored block that
intersects the selected ranges, or all blocks on the sheet when `selection` is `{ "type": "ALL" }`.
`GET_CONDITIONAL_FORMATTING` reads back authored rule families plus loaded `COLOR_SCALE_RULE`,
`DATA_BAR_RULE`, `ICON_SET_RULE`, and typed `UNSUPPORTED_RULE` entries so existing workbook
content is surfaced honestly instead of being erased at the protocol seam.

Autofilters and tables are now first-class too. `SET_AUTOFILTER` creates or replaces one
sheet-level filter range, requires a nonblank header row, and rejects overlap with any existing
table range. `CLEAR_AUTOFILTER` removes only the sheet-level filter. `SET_TABLE` creates or
replaces one workbook-global table definition, requires nonblank unique header cells, rejects
overlap with other tables, and clears any overlapping sheet-level autofilter so the table-owned
autofilter becomes authoritative on that range. Later `SET_CELL`, `SET_RANGE`, `CLEAR_RANGE`,
`APPEND_ROW`, and `APPLY_STYLE` operations that touch a table header row keep the persisted
table-column metadata converged with the visible header cells instead of letting table XML drift
until reopen.
`DELETE_TABLE` removes one existing table by name and expected sheet. `style` is either
`{ "type": "NONE" }` or a named table style with explicit stripe and emphasis flags.

`CLEAR_RANGE` removes value, style, hyperlink, and comment state from the addressed rectangle;
cells that have never been written are silently skipped. `CLEAR_HYPERLINK` and `CLEAR_COMMENT` are
no-ops when the addressed cell does not physically exist.

`SET_CELL`, `SET_RANGE`, and `APPEND_ROW` preserve any existing style, hyperlink, and comment
state already present on the targeted cells. For `DATE` and `DATE_TIME` values, GridGrind layers
the required number format onto the existing style instead of replacing fill, border, font,
alignment, or wrap state.

`APPEND_ROW` uses value-bearing row semantics: blank rows that carry only style, comment, or
hyperlink metadata do not move the append cursor, but any existing presentation or metadata state
on the reused cells is preserved when values are written there. `AUTO_SIZE_COLUMNS` uses
deterministic, content-based sizing so headless, Docker, and local runs produce the same widths.

See [docs/OPERATIONS.md](docs/OPERATIONS.md) for the full reference with all fields and examples.

---

## Reads

Reads are explicit, ordered post-mutation requests. GridGrind does not return an implicit
analysis bundle anymore; callers request exactly the workbook facts or analyses they want.

Introspection reads:
- `GET_WORKBOOK_SUMMARY`
- `GET_NAMED_RANGES`
- `GET_SHEET_SUMMARY`
- `GET_CELLS`
- `GET_WINDOW`
- `GET_MERGED_REGIONS`
- `GET_HYPERLINKS`
- `GET_COMMENTS`
- `GET_SHEET_LAYOUT`
- `GET_PRINT_LAYOUT`
- `GET_DATA_VALIDATIONS`
- `GET_CONDITIONAL_FORMATTING`
- `GET_AUTOFILTERS`
- `GET_TABLES`
- `GET_FORMULA_SURFACE`
- `GET_SHEET_SCHEMA`
- `GET_NAMED_RANGE_SURFACE`

Analysis reads:
- `ANALYZE_FORMULA_HEALTH`
- `ANALYZE_DATA_VALIDATION_HEALTH`
- `ANALYZE_CONDITIONAL_FORMATTING_HEALTH`
- `ANALYZE_AUTOFILTER_HEALTH`
- `ANALYZE_TABLE_HEALTH`
- `ANALYZE_HYPERLINK_HEALTH`
- `ANALYZE_NAMED_RANGE_HEALTH`
- `ANALYZE_WORKBOOK_FINDINGS`

Every read carries a caller-defined `requestId`, and every result echoes that `requestId` back so
agents can correlate repeated or similar reads deterministically.

`GET_WORKBOOK_SUMMARY` now returns a typed workbook summary: `kind=EMPTY` for zero-sheet
workbooks and `kind=WITH_SHEETS` when one or more sheets exist. Non-empty workbooks also report
`activeSheetName` and workbook-ordered `selectedSheetNames`. `GET_SHEET_SUMMARY` now reports
sheet `visibility` plus typed `protection` state alongside the existing row and column facts.
`GET_SHEET_LAYOUT` returns pane state, effective zoom, and per-row or per-column layout facts for
one sheet. Row and column entries now include explicit size plus `hidden`, `outlineLevel`, and
`collapsed` where Excel exposes that state. `GET_PRINT_LAYOUT` returns the supported print-area,
orientation, scaling, repeating-row, repeating-column, and plain header or footer text state for
one sheet.
String cell reads now expose `stringValue` plus optional `richText` runs in the same response
family, so no separate rich-text read call is needed.
`GET_HYPERLINKS` returns hyperlinks in the same discriminated shape used by `SET_HYPERLINK`
targets. `FILE` targets come back in the `path` field as normalized plain path strings.
`GET_DATA_VALIDATIONS` returns supported validation entries plus typed unsupported entries for the
selected ranges, preserving the normalized stored coverage ranges for each rule.
`GET_CONDITIONAL_FORMATTING` returns one factual block per stored block on the selected sheet.
Supported rule reports cover `FORMULA_RULE`, `CELL_VALUE_RULE`, `COLOR_SCALE_RULE`,
`DATA_BAR_RULE`, and `ICON_SET_RULE`, while loaded workbook content outside the supported public
model is surfaced as `UNSUPPORTED_RULE`.
`GET_AUTOFILTERS` returns one factual entry per sheet-owned or table-owned autofilter on the
requested sheet. `GET_TABLES` returns workbook-global table facts including range, header and
totals row counts, column names, style metadata, and whether the table owns an autofilter.
`ANALYZE_HYPERLINK_HEALTH` resolves relative `FILE` targets against the workbook's persisted path
when one exists; when the workbook has no filesystem location yet, relative `FILE` targets are
reported as unresolved instead of being silently treated as healthy.
`ANALYZE_DATA_VALIDATION_HEALTH` reports unsupported, broken-formula, and overlapping validation
findings. `ANALYZE_CONDITIONAL_FORMATTING_HEALTH` reports broken formulas, unsupported rule
families, priority collisions, and empty ranges. `ANALYZE_AUTOFILTER_HEALTH` reports invalid
ranges, blank headers, and sheet-versus-table ownership mismatches. `ANALYZE_TABLE_HEALTH`
reports broken ranges, overlapping tables, blank or duplicate headers, and style mismatches.
`ANALYZE_WORKBOOK_FINDINGS` now aggregates all shipped analysis families: formula, data
validation, conditional formatting, autofilter, table, hyperlink, and named-range health.

Coordinate systems are split intentionally: `address` and `range` fields use Excel A1 notation,
while fields ending in `RowIndex` or `ColumnIndex` use zero-based indexes. Validation messages now
echo both forms inline so out-of-bounds failures read like `firstRowIndex 5 (Excel row 6)` and
`firstColumnIndex 5 (Excel column F)`.

For a no-save workbook-health pass, pair `persistence.type=NONE` with
`ANALYZE_WORKBOOK_FINDINGS`. This is the default lint-style workflow:

```json
{
  "source": { "type": "NEW" },
  "persistence": { "type": "NONE" },
  "operations": [],
  "reads": [
    { "type": "ANALYZE_WORKBOOK_FINDINGS", "requestId": "lint" }
  ]
}
```

`GET_WINDOW` and `GET_SHEET_SCHEMA` require `rowCount * columnCount` ≤ 250,000. `GET_WINDOW`
additionally rejects windows that extend beyond the Excel 2007 sheet boundary
(1,048,576 rows, 16,384 columns). `GET_CELLS` rejects any address that is not valid A1 notation
or that exceeds the sheet boundary (e.g. `A1048577`, `XFE1`). These are GridGrind operational
limits to prevent out-of-memory failures during JSON serialization in bounded-heap environments.
For large sheets, use `GET_SHEET_SUMMARY` to discover the populated region and tile it with
multiple bounded window reads. See [docs/LIMITATIONS.md](docs/LIMITATIONS.md) for the full limit
reference.

The runnable example set now includes
[examples/sheet-management-request.json](examples/sheet-management-request.json) for B1 sheet
copy, active and selected sheet state, visibility, protection, and summary reads, plus
[examples/rich-text-request.json](examples/rich-text-request.json) for structured `RICH_TEXT`
authoring and read-back through existing cell introspection, plus
[examples/workbook-health-request.json](examples/workbook-health-request.json) for a no-save
analysis pass that demonstrates quoted formulas for sheet names with spaces.

---

## Example

A request that builds Alice's green coffee inventory sheet:

```json
{
  "source": { "type": "NEW" },
  "persistence": {
    "type": "SAVE_AS",
    "path": "roastery/green-coffee.xlsx"
  },
  "operations": [
    { "type": "ENSURE_SHEET", "sheetName": "Inventory" },
    {
      "type": "SET_RANGE",
      "sheetName": "Inventory",
      "range": "A1:C1",
      "rows": [[
        { "type": "TEXT", "text": "Origin" },
        { "type": "TEXT", "text": "Kilos" },
        { "type": "TEXT", "text": "Cost/kg" }
      ]]
    },
    {
      "type": "SET_RANGE",
      "sheetName": "Inventory",
      "range": "A2:C3",
      "rows": [
        [
          { "type": "TEXT",   "text": "Ethiopia Yirgacheffe" },
          { "type": "NUMBER", "number": 150 },
          { "type": "NUMBER", "number": 8.40 }
        ],
        [
          { "type": "TEXT",   "text": "Colombia Huila" },
          { "type": "NUMBER", "number": 200 },
          { "type": "NUMBER", "number": 7.80 }
        ]
      ]
    },
    {
      "type": "APPLY_STYLE",
      "sheetName": "Inventory",
      "range": "A1:C1",
      "style": {
        "alignment": {
          "horizontalAlignment": "CENTER",
          "verticalAlignment": "CENTER",
          "textRotation": 15
        },
        "font": {
          "bold": true,
          "fontName": "Aptos",
          "fontHeight": { "type": "POINTS", "points": 13 },
          "fontColor": "#FFFFFF"
        },
        "fill": {
          "pattern": "THIN_HORIZONTAL_BANDS",
          "foregroundColor": "#1F4E78",
          "backgroundColor": "#D9E2F3"
        },
        "border": {
          "all": { "style": "THIN", "color": "#FFFFFF" }
        },
        "protection": {
          "locked": true
        }
      }
    },
    { "type": "EVALUATE_FORMULAS" }
  ],
  "reads": [
    { "type": "GET_WORKBOOK_SUMMARY", "requestId": "workbook" },
    {
      "type": "GET_WINDOW",
      "requestId": "inventory-window",
      "sheetName": "Inventory",
      "topLeftAddress": "A1",
      "rowCount": 3,
      "columnCount": 3
    },
    {
      "type": "GET_SHEET_SCHEMA",
      "requestId": "inventory-schema",
      "sheetName": "Inventory",
      "topLeftAddress": "A1",
      "rowCount": 3,
      "columnCount": 3
    },
    {
      "type": "ANALYZE_WORKBOOK_FINDINGS",
      "requestId": "inventory-findings"
    }
  ]
}
```

---

## Responses

A successful response carries `"status": "SUCCESS"`, a typed `persistence` outcome, and an ordered
`reads` array. Each read result returns only the payload requested by the matching read
operation, correlated by `requestId`. Cell and window reads still carry effective cell values,
hyperlink metadata, comment metadata, and styles. Style output includes effective font name, font
size, font color, underline, strikeout, fill color, and all four border sides when those style
facts can be normalized from the workbook. Persistence is now explicit too:

- `NONE`: the workbook stayed in memory only
- `SAVE_AS`: includes both the caller-provided `requestedPath` and the actual `executionPath`
- `OVERWRITE`: includes both the original `sourcePath` token and the actual `executionPath`

Successful responses may also include a `warnings` array. GridGrind currently uses this for
request-phase formula warnings such as same-request sheet names with spaces referenced without
single quotes. Use `'Sheet Name'!A1` syntax for those formulas.

A failed response carries `"status": "ERROR"` and a structured `problem` object with:

- a stable `code` (`INVALID_CELL_ADDRESS`, `INVALID_FORMULA`, `SHEET_NOT_FOUND`, and others)
- a `category` and recommended `recovery` strategy
- a `title`, `message`, and `resolution` — all written for autonomous caller consumption
- a `context` block with stage, operation index, sheet name, address, formula, and JSON path
  coordinates so an agent can locate the exact source of failure
- a `causes` list of GridGrind-classified diagnostics, including the primary classified failure
  and any supplemental stage failures

See [docs/ERRORS.md](docs/ERRORS.md) for all problem codes and the full error model.

The CLI supports `--help` / `-h` for inline usage guidance and `--version` for the packaged
version string.

---

## Documentation

| Resource | Description |
|:---------|:------------|
| [docs/DEVELOPER.md](docs/DEVELOPER.md) | Build, architecture, coverage, and QA reference |
| [docs/OPERATIONS.md](docs/OPERATIONS.md) | Complete operation reference with all fields |
| [docs/ERRORS.md](docs/ERRORS.md) | Problem codes, categories, and error model |
| [docs/QUICK_REFERENCE.md](docs/QUICK_REFERENCE.md) | Copy-paste JSON for every operation type |
| [docs/LIMITATIONS.md](docs/LIMITATIONS.md) | All hard limits: window size, Excel maximums, memory |
| [examples/](examples/) | Runnable request files |

---

## Legal

GridGrind is MIT-licensed. Its executable JAR bundles Apache-licensed dependencies (Apache POI,
Apache Log4j Core, Jackson Databind). The bundled components retain their original Apache
License 2.0 terms.

[LICENSE](LICENSE) | [NOTICE](NOTICE) | [PATENTS.md](PATENTS.md)
