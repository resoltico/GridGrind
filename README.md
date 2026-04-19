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
docker pull ghcr.io/resoltico/gridgrind:0.48.0
```

Published GHCR release images are rebuilt from a digest-pinned Java base image and ship explicit
OCI provenance plus SBOM attestations alongside the runnable multi-arch image.

Pipe a JSON request to stdin, receive a JSON response on stdout:

```bash
echo '{"source":{"type":"NEW"},"steps":[]}' \
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
docker run --rm ghcr.io/resoltico/gridgrind:latest --print-example WORKBOOK_HEALTH
```

`--print-request-template` emits the minimal valid request. `--print-protocol-catalog` emits
machine-readable JSON describing every mutation action, inspection query, and selector shape:
required and optional fields, and the accepted shapes for polymorphic inputs. Both `--help` and
`--print-protocol-catalog` are treated as public contract surfaces and are black-box
regression-tested from the packaged JAR and Docker image, not only from source builds.
`--print-example <id>` emits one contract-owned generated example request from the shipped example
registry that also backs the committed `examples/*.json` fixtures.

### Fat JAR (requires Java 26)

Download the self-contained JAR from the
[Releases page](https://github.com/resoltico/GridGrind/releases/latest):

```bash
echo '{"source":{"type":"NEW"},"steps":[]}' \
  | java -jar gridgrind.jar

java -jar gridgrind.jar --request request.json --response response.json
java -jar gridgrind.jar --help
java -jar gridgrind.jar --license
java -jar gridgrind.jar --print-request-template
java -jar gridgrind.jar --print-protocol-catalog
java -jar gridgrind.jar --print-example ASSERTION
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

### Source-backed authored values

Text-bearing mutation fields can author values inline or load them from the execution
environment:

```json
{ "type": "INLINE", "text": "Quarterly note" }
{ "type": "UTF8_FILE", "path": "authored-inputs/note.txt" }
{ "type": "STANDARD_INPUT" }
```

Binary-bearing mutation fields use the matching binary source model:

```json
{ "type": "INLINE_BASE64", "base64Data": "SGVsbG8=" }
{ "type": "FILE", "path": "authored-inputs/payload.bin" }
{ "type": "STANDARD_INPUT" }
```

- Text sources appear inside mutation fields such as `CellInput.TEXT`, `CellInput.FORMULA`,
  rich-text runs, comments, header/footer text, validation prompts, validation error alerts,
  shape text, table comments, and chart title text.
- Binary sources appear inside picture payloads, embedded-object payloads, and embedded-object
  preview images.
- Relative `UTF8_FILE` and `FILE` paths resolve from the current working directory.
- `STANDARD_INPUT` authored values require `--request <path>` on the CLI because stdin cannot
  carry both the request JSON and the authored input bytes in one invocation.
- Source-backed loading failures surface before workbook open as `INPUT_SOURCE_NOT_FOUND`,
  `INPUT_SOURCE_UNAVAILABLE`, or `INPUT_SOURCE_IO_ERROR`, and the structured response journal
  records them under `journal.inputResolution`.

### Low-memory execution policy

`execution` is an optional top-level request object for the large-file `.xlsx` paths:

```json
{
  "execution": {
    "mode": {
      "readMode": "EVENT_READ",
      "writeMode": "STREAMING_WRITE"
    },
    "journal": {
      "level": "VERBOSE"
    },
    "calculation": {
      "strategy": {
        "type": "DO_NOT_CALCULATE"
      },
      "markRecalculateOnOpen": true
    }
  }
}
```

- Omit `execution` for the default `FULL_XSSF` read and write path with `NORMAL` journaling.
- `execution.mode.readMode: EVENT_READ` is the low-memory summary reader. It supports only
  `GET_WORKBOOK_SUMMARY` and `GET_SHEET_SUMMARY`.
- `execution.mode.writeMode: STREAMING_WRITE` is the low-memory append-oriented writer. It requires
  `source.type: NEW`, supports only `ENSURE_SHEET` and `APPEND_ROW`, and allows only
  `execution.calculation.strategy: DO_NOT_CALCULATE` with optional
  `markRecalculateOnOpen: true`.
- `execution.journal.level` defaults to `NORMAL`. Use `VERBOSE` when you want live CLI stderr
  progress plus verbose event capture in the structured response journal.
- `execution.calculation` controls immediate server-side evaluation (`EVALUATE_ALL` or
  `EVALUATE_TARGETS`), cache clearing (`CLEAR_CACHES_ONLY`), and workbook-open recalc flags
  (`markRecalculateOnOpen`).
- `STREAMING_WRITE` readback is a second request against the materialized workbook, using either
  normal `FULL_XSSF` reads or summary-only `EVENT_READ`.

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
Handwritten production Java and Kotlin sources are held to explicit-import-only policy there, so
wildcard imports fail the canonical gate instead of relying on reviewer cleanup.
The release surface also includes [`scripts/verify-cli-contract.sh`](./scripts/verify-cli-contract.sh),
which black-boxes `--help` and `--print-protocol-catalog` from the built artifact so thin
transport adapters cannot drift from the core protocol contract silently.
The committed example requests under [`examples/`](./examples/) also double as promoted
protocol-request regression inputs for drawing media, charts, pivots, conditional formatting,
low-memory modes, and OOXML package security, while the deterministic `.xlsx` round-trip verifier
asserts save-and-reopen preservation for drawings, charts, pivots, validations, tables,
autofilters, and conditional formatting.

For direct Apache POI `.xlsx` parity measurement, run:

```bash
./gradlew parity
```

That task materializes the committed XSSF golden corpus, opens it through a direct-POI oracle,
compares GridGrind against the canonical parity ledger, and fails as soon as observed parity
status drifts from the checked-in baseline.

Root `./gradlew check` stays focused on `engine`, `contract`, `executor`, `authoring-java`, and
`cli`, and now includes the executor module's parity verification. The nested Jazzer build remains
local-only on purpose and
has its own `./gradlew --project-dir jazzer check` flow for deterministic nested verification.
Use `jazzer/bin/*` for active fuzzing, replay, promotion, and other local Jazzer operator work.
GitHub Actions intentionally never runs active fuzzing, and active Jazzer harness execution now
hard-fails when `GITHUB_ACTIONS=true`. Active fuzz launcher tasks also preload a tiny
project-owned premain agent so Java 26 live fuzzing does not depend on a late external attach. The
supported wrappers also force active fuzz onto `--no-daemon` and own interrupt cleanup so canceled
local runs do not strand a live harness JVM.

## Java Authoring API

If you are already in Java, you do not need to hand-author JSON. GridGrind now ships the
`authoring-java` module as a fluent Java layer above the canonical contract and executor:

```java
import dev.erst.gridgrind.authoring.GridGrindPlan;
import dev.erst.gridgrind.authoring.Targets;
import dev.erst.gridgrind.authoring.Values;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import java.nio.file.Path;
import java.util.List;

GridGrindPlan plan =
    GridGrindPlan.newWorkbook()
        .saveAs(Path.of("budget.xlsx"))
        .journal(ExecutionJournalLevel.VERBOSE)
        .mutate(Targets.sheet("Budget").ensureExists())
        .mutate(
            Targets.range("Budget", "A1:B3")
                .setRows(
                    List.of(
                        Values.row(Values.text("Item"), Values.text("Amount")),
                        Values.row(Values.text("Hosting"), Values.number(100.0)),
                        Values.row(Values.text("Travel"), Values.number(50.0)))))
        .mutate(
            Targets.tableOnSheet("BudgetTable", "Budget")
                .define(
                    new TableInput(
                        "BudgetTable",
                        "Budget",
                        "A1:B3",
                        false,
                        new TableStyleInput.None())))
        .mutate(
            Targets.table("BudgetTable")
                .rowByKey("Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .set(Values.number(125.0)))
        .inspect(
            Targets.table("BudgetTable")
                .rowByKey("Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .read())
        .assertThat(
            Targets.table("BudgetTable")
                .rowByKey("Item", Values.textFile(Path.of("authored-inputs", "item.txt")))
                .cell("Amount")
                .valueEquals(Values.expectedNumber(125.0)));
```

That Java layer compiles down to the same canonical `WorkbookPlan` consumed by the CLI. It does
not expose engine internals, it does not create a second mutation model, and it uses the same
selectors, assertions, journaling, calculation policy, and source-backed input rules as the JSON
surface. See [examples/java-authoring-workflow.java](./examples/java-authoring-workflow.java) for
the shipped compile-verified example.

---

## The Request Contract

A request has ordered `steps[]` plus optional persistence:

**MUTATION** steps change workbook state — create sheets, fill cells, apply styles, insert rows,
build tables, set scalar formulas, and author pictures, shapes, embedded objects, supported
simple charts, or limited pivot tables.

**ASSERTION** steps verify workbook state after the earlier authored steps have run. They are
first-class, not just caller-side conventions: passed assertions are acknowledged in the success
response, and failed assertions terminate the workflow with `ASSERTION_FAILED` plus structured
observed facts. Presence-style assertions are count-based: when an exact named range, chart,
table, or pivot-table selector matches nothing, GridGrind treats that as zero observed entities
instead of leaking lower-level `*_NOT_FOUND` read failures.

**INSPECTION** steps observe the workbook after prior steps succeed — cell windows, schemas,
hyperlinks, table definitions, drawing inventories, chart metadata, pivot-table metadata, pivot
health, formula health, aggregate workbook findings, validation health, and more. Non-mutating.

**Persistence** writes the file after all authored steps succeed. Omit it or set `NONE` to keep
the workbook in memory only — useful when you only need verification or read results, or when Bob
is running his health check before deciding whether to file a report.

The whole pipeline is a complete extraction or nothing at all. `.xlsx` only; `.xls`, `.xlsm`,
and `.xlsb` are rejected.

Two contract details matter often in agent-generated requests:
- `formulaEnvironment` is optional. Use it when server-side evaluation needs external workbook
  bindings, cached-value fallback for missing external references, or template-backed UDF
  registration.
- Text and binary mutation payloads are source-backed. Use `INLINE`, `UTF8_FILE`, or
  `STANDARD_INPUT` for text-bearing fields and `INLINE_BASE64`, `FILE`, or `STANDARD_INPUT` for
  binary-bearing fields. CLI requests that use `STANDARD_INPUT` authored values must supply the
  request JSON through `--request <path>`.
- Integer-typed JSON fields are strict integers. Fractional JSON numbers are rejected instead of
  being rounded or truncated into fields such as row indexes, column indexes, or
  `sheetDefaults.defaultColumnWidth`.
- Formula payloads are scalar only. Array-formula braces such as `{=SUM(A1:A10*B1:B10)}` are
  rejected as `INVALID_FORMULA`. `LAMBDA` and `LET` are currently rejected as
  `INVALID_FORMULA` because Apache POI cannot parse them, and other newer Excel constructs may
  fail the same way. Loaded formulas that POI parses but cannot evaluate surface as
  `UNSUPPORTED_FORMULA`.
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
copy-paste snippets for every type. The public example surface is now a curated generated set that
comes from the same contract-owned registry used by `gridgrind --print-example <id>`. The
committed examples are mirror outputs of that registry:

- [examples/budget-request.json](examples/budget-request.json): selector-first budget sheet with
  styling, formula totals, workbook summary, cell readback, window readback, and schema
  inspection.
- [examples/workbook-health-request.json](examples/workbook-health-request.json): compact no-save
  health pass combining `GET_SHEET_SUMMARY`, `ANALYZE_FORMULA_HEALTH`,
  `ANALYZE_WORKBOOK_FINDINGS`, and `GET_CELLS`.
- [examples/assertion-request.json](examples/assertion-request.json): first-class
  mutate-then-verify loop with cell-value assertions, analysis-severity assertions, and verbose
  journaling.
- [examples/source-backed-input-request.json](examples/source-backed-input-request.json):
  source-backed authoring for file-loaded text, file-loaded formulas, and file-loaded binary
  payloads.
- [examples/large-file-modes-request.json](examples/large-file-modes-request.json): low-memory
  execution with top-level `execution`, append-oriented `STREAMING_WRITE`, and explicit
  `markRecalculateOnOpen=true`.
- [examples/chart-request.json](examples/chart-request.json): supported simple-chart authoring
  with named-range-backed series and factual chart readback.
- [examples/pivot-request.json](examples/pivot-request.json): pivot authoring from a contiguous
  range with `GET_PIVOT_TABLES` readback and `ANALYZE_PIVOT_TABLE_HEALTH`.
- [examples/package-security-inspect-request.json](examples/package-security-inspect-request.json):
  encrypted source open with `source.security.password` plus factual `GET_PACKAGE_SECURITY`
  inspection.
- [examples/file-hyperlink-health-request.json](examples/file-hyperlink-health-request.json):
  file and document hyperlink authoring with explicit hyperlink-health analysis.
- [examples/introspection-analysis-request.json](examples/introspection-analysis-request.json):
  batch factual reads plus formula, hyperlink, named-range, and aggregate workbook analysis in one
  ordered plan.

### Alice — building an inventory sheet

```json
{
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "SAVE_AS",
    "path": "roastery/green-coffee.xlsx"
  },
  "steps": [
    {
      "stepId": "step-01-ensure-sheet",
      "target": {
        "type": "BY_NAME",
        "name": "Inventory"
      },
      "action": {
        "type": "ENSURE_SHEET"
      }
    },
    {
      "stepId": "step-02-set-range",
      "target": {
        "type": "BY_RANGE",
        "sheetName": "Inventory",
        "range": "A1:C1"
      },
      "action": {
        "type": "SET_RANGE",
        "rows": [
          [
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Origin"
              }
            },
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Kilos"
              }
            },
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Cost/kg"
              }
            }
          ]
        ]
      }
    },
    {
      "stepId": "step-03-set-range",
      "target": {
        "type": "BY_RANGE",
        "sheetName": "Inventory",
        "range": "A2:C3"
      },
      "action": {
        "type": "SET_RANGE",
        "rows": [
          [
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Ethiopia Yirgacheffe"
              }
            },
            {
              "type": "NUMBER",
              "number": 150
            },
            {
              "type": "NUMBER",
              "number": 8.4
            }
          ],
          [
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Colombia Huila"
              }
            },
            {
              "type": "NUMBER",
              "number": 200
            },
            {
              "type": "NUMBER",
              "number": 7.8
            }
          ]
        ]
      }
    },
    {
      "stepId": "step-04-apply-style",
      "target": {
        "type": "BY_RANGE",
        "sheetName": "Inventory",
        "range": "A1:C1"
      },
      "action": {
        "type": "APPLY_STYLE",
        "style": {
          "alignment": {
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "CENTER"
          },
          "font": {
            "bold": true,
            "fontName": "Aptos",
            "fontHeight": {
              "type": "POINTS",
              "points": 13
            },
            "fontColor": "#FFFFFF"
          },
          "fill": {
            "pattern": "THIN_HORIZONTAL_BANDS",
            "foregroundColor": "#1F4E78",
            "backgroundColor": "#D9E2F3"
          },
          "border": {
            "all": {
              "style": "THIN",
              "color": "#FFFFFF"
            }
          },
          "protection": {
            "locked": true
          }
        }
      }
    },
    {
      "stepId": "workbook",
      "target": {
        "type": "CURRENT"
      },
      "query": {
        "type": "GET_WORKBOOK_SUMMARY"
      }
    },
    {
      "stepId": "inventory-window",
      "target": {
        "type": "RECTANGULAR_WINDOW",
        "sheetName": "Inventory",
        "topLeftAddress": "A1",
        "rowCount": 3,
        "columnCount": 3
      },
      "query": {
        "type": "GET_WINDOW"
      }
    },
    {
      "stepId": "inventory-findings",
      "target": {
        "type": "CURRENT"
      },
      "query": {
        "type": "ANALYZE_WORKBOOK_FINDINGS"
      }
    }
  ],
  "execution": {
    "calculation": {
      "strategy": {
        "type": "EVALUATE_ALL"
      }
    }
  }
}
```

### Bob — health check before filing a report

No persistence, no write. The workbook stays in memory, all analysis families run, findings come
back in the response.

```json
{
  "source": {
    "type": "EXISTING",
    "path": "roastery/green-coffee.xlsx"
  },
  "persistence": {
    "type": "NONE"
  },
  "steps": [
    {
      "stepId": "health",
      "target": {
        "type": "CURRENT"
      },
      "query": {
        "type": "ANALYZE_WORKBOOK_FINDINGS"
      }
    }
  ]
}
```

Response — clean workbook:

```json
{
  "status": "SUCCESS",
  "protocolVersion": "V1",
  "journal": {
    "planId": "health-check",
    "level": "NORMAL",
    "source": {
      "type": "NEW",
      "path": null
    },
    "persistence": {
      "type": "NONE",
      "path": null
    },
    "validation": {
      "status": "SUCCEEDED",
      "durationMillis": 1,
      "detail": "succeeded"
    },
    "inputResolution": {
      "status": "NOT_STARTED",
      "durationMillis": 0,
      "detail": "not started"
    },
    "open": {
      "status": "SUCCEEDED",
      "durationMillis": 2,
      "detail": "succeeded"
    },
    "calculation": {
      "preflight": {
        "status": "NOT_STARTED",
        "durationMillis": 0,
        "detail": "not started"
      },
      "execution": {
        "status": "NOT_STARTED",
        "durationMillis": 0,
        "detail": "not started"
      }
    },
    "persistencePhase": {
      "status": "NOT_STARTED",
      "durationMillis": 0,
      "detail": "not started"
    },
    "close": {
      "status": "SUCCEEDED",
      "durationMillis": 1,
      "detail": "succeeded"
    },
    "steps": [],
    "events": [],
    "outcome": {
      "status": "SUCCEEDED",
      "plannedStepCount": 1,
      "completedStepCount": 1,
      "failedStepCount": 0
    }
  },
  "persistence": {
    "type": "NONE"
  },
  "warnings": [],
  "assertions": [],
  "inspections": [
    {
      "type": "ANALYZE_WORKBOOK_FINDINGS",
      "stepId": "health",
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

A successful response includes a structured `journal`, the persistence outcome, and ordered
`assertions` and `inspections` arrays. Assertion acknowledgments preserve authored step order for
passed verification steps, and each inspection result is correlated to its originating inspection
step by `stepId`. The journal always records `validation`, `inputResolution`, `open`,
`calculation`, `persistencePhase`, and `close` as top-level phases.

A failed response still carries the structured `journal` together with a structured `problem` — a
stable code, a category and recovery strategy, a message and resolution written for autonomous
caller consumption, and a context block with step index, step id, step type, sheet name, address,
and JSON path coordinates so an agent can locate the exact source of failure. Assertion
mismatches additionally carry `problem.assertionFailure`, which includes the failed assertion
contract plus the observed factual read payloads that caused the mismatch. Source-backed input
loading failures stop the request before workbook open and surface as `INPUT_SOURCE_NOT_FOUND`,
`INPUT_SOURCE_UNAVAILABLE`, or `INPUT_SOURCE_IO_ERROR`.

Use `execution.journal.level=VERBOSE` when you want the same fine-grained phase and step events to
stream to CLI stderr while the request is still running.

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
