<!--
RETRIEVAL_HINTS:
  keywords: [gridgrind, excel, xlsx, workbook, agent, automation, json, protocol, java, apache-poi, spreadsheet, ai]
  answers: [what is gridgrind, how to use gridgrind, excel automation for ai agents, json workbook api, gridgrind quick start, gridgrind operations, gridgrind install]
  related: [docs/QUICK_REFERENCE.md, docs/OPERATIONS.md, docs/POI_EXCEL_CAPABILITY_INVENTORY.md, docs/ERRORS.md]
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

## Docs

- [docs/QUICK_REFERENCE.md](./docs/QUICK_REFERENCE.md) for copy-paste JSON snippets
- [docs/OPERATIONS.md](./docs/OPERATIONS.md) for the full request and read reference
- [docs/POI_EXCEL_CAPABILITY_INVENTORY.md](./docs/POI_EXCEL_CAPABILITY_INVENTORY.md) for the
  current public `.xlsx` capability map against Apache POI XSSF
- [docs/LIMITATIONS.md](./docs/LIMITATIONS.md) for hard ceilings and operating limits

---

## Running GridGrind

### Container (recommended)

No install required beyond Docker. Works on macOS, Linux, and Windows on both x64 and ARM:

```bash
docker pull ghcr.io/resoltico/gridgrind:latest
```

To pin to a specific release (the container registry retains the last 5 releases):

```bash
docker pull ghcr.io/resoltico/gridgrind:0.43.0
```

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
docker run --rm ghcr.io/resoltico/gridgrind:latest --license
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
java -jar gridgrind.jar --license
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

### Low-memory execution modes

`executionMode` is an optional top-level request object for the large-file `.xlsx` paths:

```json
{
  "executionMode": {
    "readMode": "EVENT_READ",
    "writeMode": "STREAMING_WRITE"
  }
}
```

- Omit `executionMode` for the default `FULL_XSSF` read and write path.
- `readMode: EVENT_READ` is the low-memory summary reader. It supports only
  `GET_WORKBOOK_SUMMARY` and `GET_SHEET_SUMMARY`.
- `writeMode: STREAMING_WRITE` is the low-memory append-oriented writer. It requires
  `source.type: NEW` and supports only `ENSURE_SHEET`, `APPEND_ROW`, and
  `FORCE_FORMULA_RECALC_ON_OPEN`.
- The two modes can be combined. `STREAMING_WRITE` can read back its materialized workbook either
  through normal `FULL_XSSF` reads or summary-only `EVENT_READ`.

---

## Developer Verification

Source builds and contributor verification use the project Gradle wrapper plus the repository root
gate:

```bash
./check.sh
```

That local whole-repo gate runs the root quality checks, nested Jazzer verification, packaging
smoke checks, and Docker smoke coverage in one supported sequence. Jazzer-specific operator flows
also remain available through the scripts under [`jazzer/bin`](./jazzer/bin/).

For direct Apache POI `.xlsx` parity measurement, run:

```bash
./gradlew parity
```

That task materializes the committed XSSF golden corpus, opens it through a direct-POI oracle,
compares GridGrind against the canonical parity ledger, and fails as soon as observed parity
status drifts from the checked-in baseline.

Root `./gradlew check` stays focused on `engine`, `protocol`, and `cli`, and now includes the
protocol module's parity verification. The nested Jazzer build remains local-only on purpose and
has its own `./gradlew --project-dir jazzer check` flow for deterministic nested verification.
Use `jazzer/bin/*` for active fuzzing, replay, promotion, and other local Jazzer operator work.
GitHub Actions intentionally never runs active fuzzing, and active Jazzer harness execution now
hard-fails when `GITHUB_ACTIONS=true`. Active fuzz launcher tasks also preload a tiny
project-owned premain agent so Java 26 live fuzzing does not depend on a late external attach. The
supported wrappers also force active fuzz onto `--no-daemon` and own interrupt cleanup so canceled
local runs do not strand a live harness JVM.

---

## The Request Contract

A request has three parts that always run in this order:

**Operations** write to the workbook — create sheets, fill cells, apply styles, insert rows, build
tables, set formulas, and author pictures, shapes, embedded objects, supported simple charts, or
limited pivot tables.
They run in sequence; the first failure stops everything.

**Reads** observe the workbook after all operations succeed — cell windows, schemas, hyperlinks,
table definitions, drawing inventories, chart metadata, pivot-table metadata, pivot health,
formula health, validation health, and more. Non-mutating.

**Persistence** writes the file after all reads succeed. Omit it or set `NONE` to keep the
workbook in memory only — useful when you only need the read results, or when Bob is running his
health check before deciding whether to file a report.

The whole pipeline is a complete extraction or nothing at all. `.xlsx` only; `.xls`, `.xlsm`,
and `.xlsb` are rejected.

Two contract details matter often in agent-generated requests:
- `formulaEnvironment` is optional. Use it when server-side evaluation needs external workbook
  bindings, cached-value fallback for missing external references, or template-backed UDF
  registration.
- `source.type: EXISTING` accepts optional `source.security.password` for encrypted OOXML
  workbooks.
- `SAVE_AS` and `OVERWRITE` accept optional `persistence.security.encryption` and
  `persistence.security.signature`. Signed workbooks that are copied without mutations keep their
  existing signatures automatically; signed workbooks that are mutated require explicit
  `persistence.security.signature` re-signing before they can be persisted.
- `GET_PACKAGE_SECURITY` returns factual OOXML package-encryption and package-signature state on
  the full-XSSF read path. `EVENT_READ` does not support it.
- `SET_TABLE.table.showTotalsRow` is optional. Omit it unless the table really includes a totals
  row; the default is `false`.
- Authored drawing mutations use explicit zero-based anchors. `SET_PICTURE`, `SET_SHAPE`,
  `SET_EMBEDDED_OBJECT`, `SET_CHART`, and `SET_DRAWING_OBJECT_ANCHOR` currently accept only
  `DrawingAnchorInput` payloads of type `TWO_CELL`.
- `SET_SHAPE` and `SET_CHART` validate before mutating. Failed authored shape or chart requests do
  not leave partial drawing artifacts behind.
- `GET_CHARTS` returns supported simple `BAR`, `LINE`, and `PIE` charts authoritatively.
  Unsupported plot families or multi-plot combinations are returned as explicit `UNSUPPORTED`
  entries with preserved plot-type tokens.
- `SET_PIVOT_TABLE` supports contiguous `RANGE`, existing `NAMED_RANGE`, and existing `TABLE`
  sources. `rowLabels`, `columnLabels`, `reportFilters`, and `dataFields` must use disjoint source
  columns because POI persists one role per pivot field, and authored report filters require
  `anchor.topLeftAddress` on Excel row `3` or lower so the page-filter layout has room above the
  rendered body.
- `GET_PIVOT_TABLES` returns supported pivots authoritatively and surfaces malformed or unsupported
  loaded pivot detail as explicit `UNSUPPORTED` entries instead of aborting the read.
- Formula-backed chart titles and series titles must resolve to one cell, either directly or
  through a defined name. Category and value series sources may still target contiguous ranges or
  defined names.
- Loaded chart metadata degrades truthfully instead of aborting sheet reads: blank stored titles
  come back as `NONE`, sparse literal caches keep missing points as empty strings, and a surviving
  orphaned graphic frame is reported as a read-only `GRAPHIC_FRAME` drawing object when its chart
  relation is gone.
- `GET_DRAWING_OBJECT_PAYLOAD` extracts binary bytes only for named pictures and embedded objects.
  Use `GET_DRAWING_OBJECTS` for picture, chart, shape, and embedded-object metadata plus factual
  one-cell, two-cell, or absolute anchors.
- Column structural edits are workbook-wide guarded. `INSERT_COLUMNS`, `DELETE_COLUMNS`, and
  `SHIFT_COLUMNS` are rejected when any formula cells or formula-defined named ranges exist
  anywhere in the workbook.

---

## Examples

Write any `.xlsx` workbook structure from a single JSON request. Read back exactly the facts you
need. Run health analysis in the same pass or on its own. See [docs/OPERATIONS.md](docs/OPERATIONS.md)
for the full field reference and [docs/QUICK_REFERENCE.md](docs/QUICK_REFERENCE.md) for
copy-paste snippets for every type. The committed
[examples/structural-layout-request.json](examples/structural-layout-request.json) example shows
panes, zoom, sheet presentation, ignored-error suppression, and advanced print layout together
with factual `GET_SHEET_LAYOUT` and `GET_PRINT_LAYOUT` readback. The committed
[examples/table-autofilter-request.json](examples/table-autofilter-request.json) example shows the
defaulted `SET_TABLE` shape with `showTotalsRow` omitted. The committed
[examples/advanced-readback-request.json](examples/advanced-readback-request.json) example shows
the richer factual readback surface for workbook protection, comment runs and anchors, advanced
print setup, structured style colors and gradients, autofilter criteria and sort state, advanced
table metadata, and workbook-health analysis. The committed
[examples/advanced-mutation-request.json](examples/advanced-mutation-request.json) example shows
the full workbook-core mutation surface for password-bearing protection, formula-defined named
ranges, advanced table and autofilter mutation, advanced conditional formatting, rich comments,
and advanced page setup. The committed
[examples/drawing-media-request.json](examples/drawing-media-request.json) example shows the
drawing and media surface: picture, shape, and embedded-object authoring, explicit anchor
replacement, drawing payload extraction, and comment coexistence on the same sheet. The committed
[examples/chart-request.json](examples/chart-request.json) example shows the chart surface:
supported `BAR` chart authoring, named-range-backed series binding, explicit chart-anchor
replacement, factual `GET_CHARTS` readback, and matching chart inventory in
`GET_DRAWING_OBJECTS`. The committed
[examples/pivot-request.json](examples/pivot-request.json) example shows the pivot surface:
range-backed, named-range-backed, and table-backed pivot authoring, factual
`GET_PIVOT_TABLES` readback, and `ANALYZE_PIVOT_TABLE_HEALTH`. The committed
[examples/package-security-create-request.json](examples/package-security-create-request.json) and
[examples/package-security-inspect-request.json](examples/package-security-inspect-request.json)
examples show the package-security surface: encrypted workbook authoring, encrypted source open
with `source.security.password`, and factual `GET_PACKAGE_SECURITY` inspection. The committed
[examples/large-file-modes-request.json](examples/large-file-modes-request.json) example shows the
low-memory execution surface: top-level `executionMode`, append-oriented `STREAMING_WRITE`, and
summary-only `EVENT_READ` readback over the materialized workbook. The
committed
[examples/formula-environment-request.json](examples/formula-environment-request.json) example
shows the formula surface: top-level `formulaEnvironment`, template-backed UDF registration,
targeted formula evaluation, and explicit formula-cache clearing.

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
| [docs/DEVELOPER_GRADLE.md](docs/DEVELOPER_GRADLE.md) | Gradle system map, setup rationale, and review checklist |
| [examples/](examples/) | Runnable request files |

---

## Legal

GridGrind is MIT-licensed. Its executable JAR bundles Apache POI, Jackson, Apache Log4j, Apache
Commons, Apache XMLBeans, SparseBitSet (all Apache 2.0), and CurvesAPI (BSD 3-Clause). See
[NOTICE](NOTICE) for the complete attribution list and [PATENTS.md](PATENTS.md) for patent
considerations.

[LICENSE](LICENSE) | [NOTICE](NOTICE) | [PATENTS.md](PATENTS.md)
