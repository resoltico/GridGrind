<!--
RETRIEVAL_HINTS:
  keywords: [gridgrind, excel, xlsx, workbook, agent, automation, json, protocol, java, apache-poi, spreadsheet, ai]
  answers: [what is gridgrind, how to use gridgrind, excel automation for ai agents, json workbook api, gridgrind quick start, gridgrind operations, gridgrind install]
  related: [docs/QUICK_REFERENCE.md, docs/OPERATIONS.md, docs/ERRORS.md]
-->

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
docker pull ghcr.io/resoltico/gridgrind:0.14.0
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

```bash
echo '{"source":{"type":"NEW"},"operations":[],"reads":[]}' \
  | java -jar gridgrind.jar

java -jar gridgrind.jar --request request.json --response response.json
java -jar gridgrind.jar --version
java -jar gridgrind.jar --print-request-template
java -jar gridgrind.jar --print-protocol-catalog
```

---

## How It Works

GridGrind currently supports `.xlsx` workbooks only. Malformed JSON is rejected as
`INVALID_JSON`. Syntactically valid JSON that does not match the protocol model is rejected as
`INVALID_REQUEST_SHAPE`. Parsed requests with invalid business data, such as `.xls`, `.xlsm`,
`.xlsb`, or other non-`.xlsx` workbook paths, are rejected as `INVALID_REQUEST`.

Every request follows the same pipeline:

1. **Operations** run in order — create sheets, write cells, apply styles, evaluate formulas.
2. **Reads** run after all operations succeed — explicit post-mutation introspection and analysis
   operations return only the workbook facts the caller requested.
3. **Persistence** happens last — the workbook is written only after reads succeed.

If any step fails, GridGrind returns a structured error and no file is written. That gives
autonomous callers deterministic failure semantics instead of "error after side effect."

---

## Operations

| Operation | What It Does |
|:----------|:-------------|
| `ENSURE_SHEET` | Create a sheet if it does not already exist |
| `RENAME_SHEET` | Rename an existing sheet to a new name |
| `DELETE_SHEET` | Remove an existing sheet |
| `MOVE_SHEET` | Move an existing sheet to a zero-based position |
| `MERGE_CELLS` | Merge a rectangular A1-style range |
| `UNMERGE_CELLS` | Remove a merged region by exact range match |
| `SET_COLUMN_WIDTH` | Set one or more column widths using character units |
| `SET_ROW_HEIGHT` | Set one or more row heights using point units |
| `FREEZE_PANES` | Freeze panes using explicit split and visible-origin coordinates |
| `SET_CELL` | Write a typed value to a single cell |
| `SET_RANGE` | Write a rectangular grid of typed values |
| `CLEAR_RANGE` | Remove all values and styles from a rectangular range |
| `SET_HYPERLINK` | Attach a hyperlink to a single cell |
| `CLEAR_HYPERLINK` | Remove a hyperlink from a cell; no-op when the cell does not exist |
| `SET_COMMENT` | Attach a plain-text comment to a single cell |
| `CLEAR_COMMENT` | Remove a comment from a cell; no-op when the cell does not exist |
| `APPLY_STYLE` | Apply number formats, font styling, fills, borders, alignment, or wrap to a range |
| `SET_NAMED_RANGE` | Create or replace one workbook- or sheet-scoped named range |
| `DELETE_NAMED_RANGE` | Remove one existing workbook- or sheet-scoped named range |
| `APPEND_ROW` | Append a row of typed values after the last populated row |
| `AUTO_SIZE_COLUMNS` | Size columns to fit their contents |
| `EVALUATE_FORMULAS` | Force formula recalculation before save |
| `FORCE_FORMULA_RECALCULATION_ON_OPEN` | Mark the workbook to recalculate when opened in Excel |

Cell values accept types: `TEXT`, `NUMBER`, `BOOLEAN`, `FORMULA`, `DATE`, `DATE_TIME`, `BLANK`.

Sheet-management operations use strict semantics: missing sheets fail, `MOVE_SHEET.targetIndex`
is zero-based, and `RENAME_SHEET` requires a valid destination name that does not conflict with
another sheet.

Structural layout operations are also strict: `MERGE_CELLS` uses A1-style ranges, `UNMERGE_CELLS`
requires an exact merged-region match, `SET_COLUMN_WIDTH.widthCharacters` is converted to POI
width units with `round(widthCharacters * 256)`, `SET_ROW_HEIGHT.heightPoints` is passed through
as Excel point units, and `FREEZE_PANES` uses explicit split plus visible-origin coordinates.

`APPLY_STYLE` supports number formats, bold/italic, wrap, horizontal and vertical alignment,
`fontName`, typed `fontHeight`, `fontColor`, `underline`, `strikeout`, `fillColor`, and per-side
border styles through a nested `border` patch. `fontHeight` accepts either point units or exact
twips. Style analysis reports `fontHeight.twips` and `fontHeight.points` alongside the other
effective font, fill, and border facts.

Authoring metadata is also supported directly. `SET_HYPERLINK` and `CLEAR_HYPERLINK`
work on one cell at a time with typed `URL`, `EMAIL`, `FILE`, or `DOCUMENT` targets.
`SET_COMMENT` and `CLEAR_COMMENT` work on one cell at a time with plain-text comment content,
author, and visible state. `SET_NAMED_RANGE` and `DELETE_NAMED_RANGE` work on workbook or sheet
scope with explicit sheet-qualified cell or range targets. Analysis can now return cell
`hyperlink()` and `comment()` metadata plus workbook-level `namedRanges`.

`CLEAR_RANGE` removes value, style, hyperlink, and comment state from the addressed rectangle;
cells that have never been written are silently skipped. `CLEAR_HYPERLINK` and `CLEAR_COMMENT` are
no-ops when the addressed cell does not physically exist.

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
- `GET_FORMULA_SURFACE`
- `GET_SHEET_SCHEMA`
- `GET_NAMED_RANGE_SURFACE`

Analysis reads:
- `ANALYZE_FORMULA_HEALTH`
- `ANALYZE_HYPERLINK_HEALTH`
- `ANALYZE_NAMED_RANGE_HEALTH`
- `ANALYZE_WORKBOOK_FINDINGS`

Every read carries a caller-defined `requestId`, and every result echoes that `requestId` back so
agents can correlate repeated or similar reads deterministically.

`GET_WINDOW` and `GET_SHEET_SCHEMA` require `rowCount * columnCount` ≤ 250,000. `GET_WINDOW`
additionally rejects windows that extend beyond the Excel 2007 sheet boundary
(1,048,576 rows, 16,384 columns). `GET_CELLS` rejects any address that is not valid A1 notation
or that exceeds the sheet boundary (e.g. `A1048577`, `XFE1`). These are GridGrind operational
limits to prevent out-of-memory failures during JSON serialization in bounded-heap environments.
For large sheets, use `GET_SHEET_SUMMARY` to discover the populated region and tile it with
multiple bounded window reads. See [docs/LIMITATIONS.md](docs/LIMITATIONS.md) for the full limit
reference.

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
        "bold": true,
        "fontName": "Aptos",
        "fontHeight": { "type": "POINTS", "points": 13 },
        "fontColor": "#FFFFFF",
        "fillColor": "#1F4E78",
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "CENTER",
        "border": {
          "all": { "style": "THIN" }
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

A failed response carries `"status": "ERROR"` and a structured `problem` object with:

- a stable `code` (`INVALID_CELL_ADDRESS`, `INVALID_FORMULA`, `SHEET_NOT_FOUND`, and others)
- a `category` and recommended `recovery` strategy
- a `title`, `message`, and `resolution` — all written for autonomous caller consumption
- a `context` block with stage, operation index, sheet name, address, formula, and JSON path
  coordinates so an agent can locate the exact source of failure
- a `causes` chain for inspecting the underlying exception family

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
