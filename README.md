[![GridGrind Art](https://raw.githubusercontent.com/resoltico/GridGrind/main/images/GridGrind.jpg)](https://github.com/resoltico/GridGrind)

[![CI](https://github.com/resoltico/GridGrind/actions/workflows/ci.yml/badge.svg)](https://github.com/resoltico/GridGrind/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/resoltico/GridGrind)](https://github.com/resoltico/GridGrind/releases/latest)

# GridGrind — .xlsx workbook automation from a JSON request

GridGrind is a `.xlsx` automation engine. Describe workbook work as a JSON request — create sheets, write cells, build tables, assert results, read facts back. GridGrind runs the whole plan and returns a structured JSON response. If anything fails, nothing is written. No Excel installation required.

The common alternative is a mix of file-manipulation libraries, hand-written scripts, and post-write checks that run after the file is already saved — with no clean rollback when something fails mid-run. GridGrind replaces that split with one atomic pass: request in, result out, workbook saved only when every step succeeds.

- Write `.xlsx` workbooks from JSON: sheets, cells, styles, tables, formulas, charts, drawings
- Read facts back in the same plan: cell values, sheet layout, health analysis, pivot data
- Assert workbook state mid-run — a failed assertion stops the plan before saving
- Run from Docker or a self-contained JAR, against new workbooks or existing `.xlsx` files

[First run →](docs/QUICK_START.md) · [Snippets](docs/QUICK_REFERENCE.md) · [Latest release](https://github.com/resoltico/GridGrind/releases/latest)

## One request, one result

Alice logs green coffee arrivals at High Ground Roasters every Monday. One JSON request writes new lot data, asserts the workbook is clean, reads facts back, and saves — or stops cleanly if any step fails:

```json
{
  "protocolVersion": "V1",
  "source": { "type": "NEW" },
  "persistence": { "type": "SAVE_AS", "path": "lots.xlsx" },
  "execution": {
    "mode": { "readMode": "FULL_XSSF", "writeMode": "FULL_XSSF" },
    "journal": { "level": "NORMAL" },
    "calculation": {
      "strategy": { "type": "DO_NOT_CALCULATE" },
      "markRecalculateOnOpen": false
    }
  },
  "formulaEnvironment": {
    "externalWorkbooks": [],
    "missingWorkbookPolicy": "ERROR",
    "udfToolpacks": []
  },
  "steps": [
    {
      "stepId": "sheet",
      "target": { "type": "SHEET_BY_NAME", "name": "Lots" },
      "action": { "type": "ENSURE_SHEET" }
    },
    {
      "stepId": "log-lot",
      "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Lots", "address": "A1" },
      "action": {
        "type": "SET_CELL",
        "value": { "type": "TEXT", "source": { "type": "INLINE", "text": "Ethiopia Yirgacheffe" } }
      }
    },
    {
      "stepId": "check",
      "target": { "type": "WORKBOOK_CURRENT" },
      "assertion": {
        "type": "EXPECT_ANALYSIS_FINDING_ABSENT",
        "query": { "type": "ANALYZE_WORKBOOK_FINDINGS" },
        "code": "HYPERLINK_MALFORMED_TARGET"
      }
    },
    {
      "stepId": "read-back",
      "target": { "type": "CELL_BY_ADDRESSES", "sheetName": "Lots", "addresses": ["A1"] },
      "query": { "type": "GET_CELLS" }
    }
  ]
}
```

Write, assert, read — one plan. GridGrind saves the file only if every step, including the assertion, succeeds.
`--print-request-template` emits the canonical minimal request scaffold when you want to start from
the exact current wire envelope instead of a hand-written example.

## Where it fits

Good fit:
- Recurring `.xlsx` workbook jobs that should run the same way each time — filing, updating, checking, extracting
- Automation and agent pipelines that create or maintain Excel files without a UI
- Workbook health checks and fact extraction without saving (`"persistence": {"type": "NONE"}`)
- Environments without Excel — Linux containers, CI pipelines, server-side workflows

Skip it when:
- You need `.xls`, `.xlsm`, or `.xlsb` — GridGrind handles `.xlsx` only
- Your work is truly one-off and hand-writing JSON adds more friction than it saves
- You need interactive formula recalculation during editing (GridGrind evaluates on request)

Server-side formula behavior lives under `execution.calculation`. For append-oriented or
streaming-friendly plans, keep `strategy=DO_NOT_CALCULATE` and set
`markRecalculateOnOpen=true` when you want Excel to refresh formulas the next time a person opens
the workbook. In `STREAMING_WRITE`, the supported mutation set stays intentionally narrow:
`ENSURE_SHEET` creates tabs and `APPEND_ROW` streams data into them.

## Get it

**Docker** (recommended — no install beyond Docker, works on macOS, Linux, Windows, x64 and ARM):

```bash
docker run --pull=always --rm ghcr.io/resoltico/gridgrind:latest --help
```

**Release JAR** (requires Java 26) — download from [Releases](https://github.com/resoltico/GridGrind/releases/latest), then:

```bash
java -jar gridgrind.jar --print-example BUDGET --response budget-request.json
java -jar gridgrind.jar --request budget-request.json --response response.json
```

The artifact is self-describing: `--print-protocol-catalog` emits machine-readable JSON for every operation, `--print-example <id>` generates a ready-to-run request, and `--doctor-request` validates a request without touching a workbook.
Printed examples, discovery output, and normal execution responses all omit absent optional fields
instead of publishing explicit JSON `null` placeholders.

- [First run guide](docs/QUICK_START.md) — Docker or JAR, first successful run, common pitfalls
- [Snippets](docs/QUICK_REFERENCE.md) — request shapes, step patterns, assertion and inspection snippets
- [Java authoring](docs/JAVA_AUTHORING.md) — build the same plan from Java instead of hand-writing JSON
- [Operations reference](docs/OPERATIONS.md) — full field list for mutations, assertions, and inspections
- [Examples](examples/) — ready-to-run request files
  Includes [chart-request.json](examples/chart-request.json),
  [signature-line-request.json](examples/signature-line-request.json), and
  [large-file-modes-request.json](examples/large-file-modes-request.json)

## What you can check

- Published on [GitHub Releases](https://github.com/resoltico/GridGrind/releases) and [GHCR](https://github.com/resoltico/GridGrind/pkgs/container/gridgrind) as a multi-arch Docker image
- Built on [Apache POI XSSF](https://poi.apache.org/), a long-established Java `.xlsx` library
- Atomic writes: a failed run never produces a partial `.xlsx` file
- Help output and the protocol catalog are regression-tested from the packaged artifact, not only from source builds
- MIT-licensed, with runnable example files and a self-describing protocol catalog

## Legal

GridGrind is MIT-licensed. Its executable JAR bundles Apache POI, Jackson, Apache Log4j, Apache Commons, Apache XMLBeans, SparseBitSet (Apache 2.0), and CurvesAPI (BSD 3-Clause). See [NOTICE](NOTICE) for the complete attribution list and [PATENTS.md](PATENTS.md) for patent considerations.

[LICENSE](LICENSE) | [NOTICE](NOTICE) | [PATENTS.md](PATENTS.md)
