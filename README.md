<!--
RETRIEVAL_HINTS:
  keywords: [gridgrind, excel, xlsx, workbook, agent, automation, json, protocol, java, apache-poi, spreadsheet, ai]
  answers: [what is gridgrind, how to use gridgrind, excel automation for ai agents, json workbook api, gridgrind quick start, gridgrind operations, gridgrind install]
  related: [docs/QUICK_REFERENCE.md, docs/OPERATIONS.md, docs/ERRORS.md]
-->

[![GridGrind Art](https://raw.githubusercontent.com/resoltico/GridGrind/main/images/GridGrind.jpg)](https://github.com/resoltico/GridGrind)

[![CI](https://github.com/resoltico/GridGrind/actions/workflows/ci.yml/badge.svg)](https://github.com/resoltico/GridGrind/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/resoltico/GridGrind)](https://github.com/resoltico/GridGrind/releases/latest)

-----

# GridGrind

**A fresh grind for every workbook.**

GridGrind is a `.xlsx` engine with a JSON protocol. Send a request describing what to write, what
to read back, and where to save. GridGrind executes it and returns a structured response. If
anything fails, nothing is written — no partial files, no side effects before the error.
No Excel installation. Any caller that can send JSON can use it.

---

Alice runs sourcing at High Ground Roasters. Her books live in Excel: green coffee lots by origin,
batch yields, weekly price tiers. Every Monday her AI agent logs new arrivals, recalculates costs,
and files the purchase summary — all from a JSON request. Bob handles the financial side: his agent
reads those same workbooks, pulls the margin figures, and runs a health check before filing any
report. Neither of them touches a cell. GridGrind handles the grind.

---

## Running GridGrind

### Container (recommended)

No install required beyond Docker. Works on macOS, Linux, and Windows on both x64 and ARM:

```bash
docker pull ghcr.io/resoltico/gridgrind:latest
```

To pin to a specific release:

```bash
docker pull ghcr.io/resoltico/gridgrind:0.32.0
```

The container registry retains the last 5 releases. For older versions, download the fat JAR
from the [Releases page](https://github.com/resoltico/GridGrind/releases).

Pipe a JSON request to stdin, receive a JSON response on stdout:

```bash
echo '{"source":{"type":"NEW"},"operations":[],"reads":[]}' \
  | docker run -i ghcr.io/resoltico/gridgrind:latest
```

Mount a working directory to read request files and write `.xlsx` files back to the host:

```bash
docker run -i \
  -v "$(pwd)":/workdir \
  -w /workdir \
  ghcr.io/resoltico/gridgrind:latest \
  --request request.json \
  --response response.json
```

The artifact describes itself:

```bash
docker run --rm ghcr.io/resoltico/gridgrind:latest --help
docker run --rm ghcr.io/resoltico/gridgrind:latest --version
docker run --rm ghcr.io/resoltico/gridgrind:latest --print-request-template
docker run --rm ghcr.io/resoltico/gridgrind:latest --print-protocol-catalog
```

`--print-request-template` emits the minimal valid request. `--print-protocol-catalog` emits
machine-readable JSON describing every operation and read: required and optional fields, and the
accepted shapes for polymorphic inputs.

### Fat JAR (requires Java 26)

Download the self-contained JAR from the
[Releases page](https://github.com/resoltico/GridGrind/releases/latest):

```bash
echo '{"source":{"type":"NEW"},"operations":[],"reads":[]}' \
  | java -jar gridgrind.jar

java -jar gridgrind.jar --request request.json --response response.json
java -jar gridgrind.jar --help
java -jar gridgrind.jar --print-request-template
java -jar gridgrind.jar --print-protocol-catalog
```

`java -jar gridgrind.jar` uses the ambient shell `java`. Ensure it resolves to Java 26; see
[docs/DEVELOPER_JAVA.md](./docs/DEVELOPER_JAVA.md).

### File and path model

| | Behavior |
|:--|:--|
| No `--request` flag | Read the JSON request from stdin |
| `--request <path>` | Read the JSON request from that file |
| No `--response` flag | Write the JSON response to stdout |
| `--response <path>` | Write the JSON response to that file; parent directories are created |
| `source.path` | Open an existing workbook from that path |
| `persistence: SAVE_AS` | Write to `path`; parent directories are created |
| `persistence: OVERWRITE` | Write back to `source.path`; no path field |

Relative paths in `--request`, `--response`, `source.path`, and `persistence.path` resolve from
the current working directory. In Docker, set `-w` to the mount point so relative paths resolve
inside the mounted directory.

---

## The Request Contract

A request has three parts that always run in this order:

**Operations** write to the workbook — create sheets, fill cells, apply styles, insert rows, build
tables, set formulas. They run in sequence; the first failure stops everything.

**Reads** observe the workbook after all operations succeed — cell windows, schemas, hyperlinks,
table definitions, formula health, validation health, and more. Non-mutating.

**Persistence** writes the file after all reads succeed. Omit it or set `NONE` to keep the
workbook in memory only — useful when you only need the read results, or when Bob is running his
health check before deciding whether to file a report.

The whole pipeline is a complete extraction or nothing at all. `.xlsx` only; `.xls`, `.xlsm`,
and `.xlsb` are rejected.

---

## Examples

Write any `.xlsx` workbook structure from a single JSON request. Read back exactly the facts you
need. Run health analysis in the same pass or on its own. See [docs/OPERATIONS.md](docs/OPERATIONS.md)
for the full field reference and [docs/QUICK_REFERENCE.md](docs/QUICK_REFERENCE.md) for
copy-paste snippets for every type.

### Alice — building an inventory sheet

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
        "alignment": { "horizontalAlignment": "CENTER", "verticalAlignment": "CENTER" },
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
        "border": { "all": { "style": "THIN", "color": "#FFFFFF" } },
        "protection": { "locked": true }
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
    { "type": "ANALYZE_WORKBOOK_FINDINGS", "requestId": "inventory-findings" }
  ]
}
```

### Bob — health check before filing a report

No persistence, no write. The workbook stays in memory, all analysis families run, findings come
back in the response.

```json
{
  "source": { "type": "EXISTING", "path": "roastery/green-coffee.xlsx" },
  "persistence": { "type": "NONE" },
  "operations": [],
  "reads": [
    { "type": "ANALYZE_WORKBOOK_FINDINGS", "requestId": "health" }
  ]
}
```

Response — clean workbook:

```json
{
  "status": "SUCCESS",
  "protocolVersion": "V1",
  "persistence": { "type": "NONE" },
  "warnings": [],
  "reads": [
    {
      "type": "ANALYZE_WORKBOOK_FINDINGS",
      "requestId": "health",
      "analysis": {
        "summary": {
          "totalCount": 0,
          "errorCount": 0,
          "warningCount": 0,
          "infoCount": 0
        },
        "findings": []
      }
    }
  ]
}
```

---

## Responses

Every response carries `"status": "SUCCESS"` or `"status": "ERROR"`.

A successful response includes the persistence outcome and an ordered `reads` array where each
result is correlated to its read by `requestId`.

A failed response carries a structured `problem` — a stable code, a category and recovery
strategy, a message and resolution written for autonomous caller consumption, and a context block
with stage, operation index, sheet name, address, and JSON path coordinates so an agent can
locate the exact source of failure.

See [docs/ERRORS.md](docs/ERRORS.md) for all problem codes and the full error model.

---

## Documentation

| Resource | Description |
|:---------|:------------|
| [docs/OPERATIONS.md](docs/OPERATIONS.md) | Complete operation and read reference with all fields |
| [docs/QUICK_REFERENCE.md](docs/QUICK_REFERENCE.md) | Copy-paste JSON for every operation type |
| [docs/ERRORS.md](docs/ERRORS.md) | Problem codes, categories, and the full error model |
| [docs/LIMITATIONS.md](docs/LIMITATIONS.md) | All hard limits: window size, Excel maximums, memory |
| [docs/DEVELOPER.md](docs/DEVELOPER.md) | Build, architecture, coverage, and QA reference |
| [examples/](examples/) | Runnable request files |

---

## Legal

GridGrind is MIT-licensed. Its executable JAR bundles Apache POI, Jackson, Apache Log4j, Apache
Commons, Apache XMLBeans, SparseBitSet (all Apache 2.0), and CurvesAPI (BSD 3-Clause). See
[NOTICE](NOTICE) for the complete attribution list and [PATENTS.md](PATENTS.md) for patent
considerations.

[LICENSE](LICENSE) | [NOTICE](NOTICE) | [PATENTS.md](PATENTS.md)
