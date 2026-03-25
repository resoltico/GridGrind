<!--
RETRIEVAL_HINTS:
  keywords: [gridgrind, excel, xlsx, workbook, agent, automation, json, protocol, java, apache-poi, spreadsheet, ai]
  answers: [what is gridgrind, how to use gridgrind, excel automation for ai agents, json workbook api, gridgrind quick start, gridgrind operations, gridgrind install]
  related: [docs/QUICK_REFERENCE.md, docs/OPERATIONS.md, docs/ERRORS.md]
-->

# GridGrind

**A fresh grind for every workbook.**

GridGrind gives AI agents a typed, structured, deterministic way to create, edit, calculate,
inspect, and save `.xlsx` workbooks. Send a JSON request. Get a JSON response. No low-level
library calls, no ad hoc scripting, no brittle string parsing.

---

Alice runs sourcing at High Ground Roasters. Her business lives in Excel: green coffee
inventory by origin, roast batch yields, wholesale price tiers, weekly revenue by SKU. Every
Monday, her AI agent updates those sheets — new arrivals logged, prices recalculated, batch
reports drafted and filed. The agent does not write code to manipulate cells. It sends GridGrind
a JSON request. High Ground's numbers are always current, always correct, always where they need
to be.

---

## Running GridGrind

### Container (recommended)

No install required beyond Docker. Works on macOS, Linux, and Windows across both x64 and ARM:

```bash
docker pull ghcr.io/resoltico/gridgrind:latest
```

Pipe a JSON request to stdin, receive a JSON response on stdout:

```bash
echo '{"source":{"mode":"NEW"},"operations":[],"analysis":{"sheets":[]}}' \
  | docker run -i ghcr.io/resoltico/gridgrind:latest
```

To check the version:

```bash
docker run --rm ghcr.io/resoltico/gridgrind:latest --version
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

File paths in `persistence.path` are resolved relative to the container working directory
(`/workdir` above), so the generated `.xlsx` lands in `$(pwd)` on the host.

### Fat JAR (requires Java 26)

Download the self-contained JAR from the
[Releases page](https://github.com/resoltico/GridGrind/releases/latest) and run it the same
way as the container — stdin/stdout or explicit file paths:

```bash
echo '{"source":{"mode":"NEW"},"operations":[],"analysis":{"sheets":[]}}' \
  | java -jar gridgrind.jar

java -jar gridgrind.jar --request request.json --response response.json
java -jar gridgrind.jar --version
```

---

## How It Works

GridGrind currently supports `.xlsx` workbooks only. Requests that use `.xls`, `.xlsm`, `.xlsb`,
or other non-`.xlsx` workbook paths are rejected as `INVALID_REQUEST`.

Every request follows the same pipeline:

1. **Operations** run in order — create sheets, write cells, apply styles, evaluate formulas.
2. **Analysis** runs after all operations succeed — inspect sheets, read cell values and styles.
3. **Persistence** happens last — the workbook is written only after analysis succeeds.

If any step fails, GridGrind returns a structured error and no file is written. Agents get
deterministic failure semantics instead of "error after side effect."

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
| `APPLY_STYLE` | Apply number format, bold, italic, alignment, or wrap to a range |
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

See [docs/OPERATIONS.md](docs/OPERATIONS.md) for the full reference with all fields and examples.

---

## Example

A request that builds Alice's green coffee inventory sheet:

```json
{
  "source": { "mode": "NEW" },
  "persistence": {
    "mode": "SAVE_AS",
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
      "style": { "bold": true }
    },
    { "type": "EVALUATE_FORMULAS" }
  ],
  "analysis": {
    "sheets": [
      { "sheetName": "Inventory", "previewRowCount": 3, "previewColumnCount": 3 }
    ]
  }
}
```

---

## Responses

A successful response carries `"status": "SUCCESS"` and a structured workbook summary including
effective cell values and styles for every cell requested in `analysis`.

A failed response carries `"status": "ERROR"` and a structured `problem` object with:

- a stable `code` (`INVALID_CELL_ADDRESS`, `INVALID_FORMULA`, `SHEET_NOT_FOUND`, and others)
- a `category` and recommended `recovery` strategy
- a `title`, `message`, and `resolution` — all written for agent consumption
- a `context` block with stage, operation index, sheet name, address, formula, and JSON path
  coordinates so an agent can locate the exact source of failure
- a `causes` chain for inspecting the underlying exception family

See [docs/ERRORS.md](docs/ERRORS.md) for all problem codes and the full error model.

---

## Documentation

| Resource | Description |
|:---------|:------------|
| [docs/DEVELOPER.md](docs/DEVELOPER.md) | Build, architecture, coverage, and QA reference |
| [docs/OPERATIONS.md](docs/OPERATIONS.md) | Complete operation reference with all fields |
| [docs/ERRORS.md](docs/ERRORS.md) | Problem codes, categories, and error model |
| [docs/QUICK_REFERENCE.md](docs/QUICK_REFERENCE.md) | Copy-paste JSON for every operation type |
| [examples/](examples/) | Runnable request files |

---

## Legal

GridGrind is MIT-licensed. Its executable JAR bundles Apache-licensed dependencies (Apache POI,
Apache Log4j Core, Jackson Databind). The bundled components retain their original Apache
License 2.0 terms.

[LICENSE](LICENSE) | [NOTICE](NOTICE) | [PATENTS.md](PATENTS.md)
