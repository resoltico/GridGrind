[![GridGrind Art](https://raw.githubusercontent.com/resoltico/GridGrind/main/images/GridGrind.jpg)](https://github.com/resoltico/GridGrind)

[![Release](https://img.shields.io/github/v/release/resoltico/GridGrind?label=release)](https://github.com/resoltico/GridGrind/releases)
[![CI](https://github.com/resoltico/GridGrind/actions/workflows/ci.yml/badge.svg)](https://github.com/resoltico/GridGrind/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Java 26](https://img.shields.io/badge/java-26-orange.svg)](https://openjdk.org/projects/jdk/26/)

# GridGrind — .xlsx workbook automation from a JSON request

GridGrind is a `.xlsx` automation engine. Describe workbook work as a JSON request — create sheets,
write cells, build tables, assert results, read facts back. GridGrind runs the whole plan and
returns a structured JSON response. If anything fails, nothing is written.

The usual alternative is a mix of libraries, helper scripts, and post-write checks that run after
the file is already saved — with no clean rollback when something fails mid-run. GridGrind replaces
that split with one atomic pass: request in, result out, workbook written only when every step
succeeds.

- Write `.xlsx` workbooks from JSON: sheets, cells, styles, tables, formulas, charts, drawings
- Read facts back in the same plan: cell values, sheet layout, health analysis, pivot data
- Assert workbook state mid-run — a failed assertion stops the plan before saving
- Run from Docker or a self-contained JAR, against new workbooks or existing `.xlsx` files

[First run →](docs/QUICK_START.md) · [Snippets](docs/QUICK_REFERENCE.md)

## One request, one result

Alice logs green coffee arrivals at High Ground Roasters every Monday. One JSON request writes new
lot data, asserts the workbook is clean, reads facts back, and saves — or stops cleanly if any step
fails:

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

Write, assert, read — one plan. GridGrind saves the file only if every step, including the
assertion, succeeds.

## Where it fits

Good fit:
- Recurring `.xlsx` workbook jobs that should run the same way each time — filing, updating,
  checking, extracting
- Automation and agent pipelines that create or maintain Excel files without a UI
- Workbook health checks and fact extraction without saving (`"persistence": {"type": "NONE"}`)
- Environments without Excel — Linux containers, CI pipelines, server-side workflows
- Low-memory append pipelines — `STREAMING_WRITE` mode limits mutations to `ENSURE_SHEET`
  and `APPEND_ROW`; set `markRecalculateOnOpen=true` when Excel clients should recalculate on open

Skip it when:
- You need `.xls`, `.xlsm`, or `.xlsb` — GridGrind handles `.xlsx` only
- Your work is truly one-off and hand-writing JSON adds more friction than it saves
- You need interactive formula recalculation during editing (GridGrind evaluates on request)

## Get it

**Docker** (recommended — no install beyond Docker, works on macOS, Linux, Windows, x64 and ARM):

```bash
docker run --pull=always --rm ghcr.io/resoltico/gridgrind:latest --help
```

**Release JAR** (requires Java 26) — download from
[Releases](https://github.com/resoltico/GridGrind/releases/latest), then:

Don't know the protocol yet? Describe your workbook goal — GridGrind matches it to one of eight
task recipes and returns a runnable starter request:

```bash
java -jar gridgrind.jar --print-goal-plan "weekly coffee lot tracker with totals"
java -jar gridgrind.jar --print-task-catalog
```

Or copy one of 14 built-in working examples and run it directly:

```bash
java -jar gridgrind.jar --print-example BUDGET --response budget-request.json
java -jar gridgrind.jar --request budget-request.json --response response.json
```

- [First run guide](docs/QUICK_START.md) — Docker or JAR, first successful run, common pitfalls
- [Snippets](docs/QUICK_REFERENCE.md) — request shapes, step patterns, assertion and inspection snippets
- [Java authoring](docs/JAVA_AUTHORING.md) — build the same plan from Java instead of hand-writing JSON
- [Operations reference](docs/OPERATIONS.md) — full field list for mutations, assertions, and inspections
- [Examples](examples/) — ready-to-run request files including
  [chart-request.json](examples/chart-request.json),
  [signature-line-request.json](examples/signature-line-request.json), and
  [large-file-modes-request.json](examples/large-file-modes-request.json)

## Legal

GridGrind is MIT-licensed. Its executable JAR bundles Apache POI, Jackson, Apache Log4j, Apache
Commons, Apache XMLBeans, SparseBitSet (Apache 2.0), and CurvesAPI (BSD 3-Clause). See
[NOTICE](NOTICE) for the complete attribution list and [PATENTS.md](PATENTS.md) for patent
considerations.

[LICENSE](LICENSE) | [NOTICE](NOTICE) | [PATENTS.md](PATENTS.md) | [LICENSE-APACHE-2.0](LICENSE-APACHE-2.0) | [LICENSE-BSD-3-CLAUSE](LICENSE-BSD-3-CLAUSE)
