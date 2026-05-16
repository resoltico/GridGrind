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

## Where it fits

Good fit:
- Recurring `.xlsx` workbook jobs that should run the same way each time — filing, updating,
  checking, extracting
- Automation and agent pipelines that create or maintain Excel files without a UI
- Workbook health checks and fact extraction without saving a file
- Environments without Excel — Linux containers, CI pipelines, server-side workflows

Skip it when:
- You need `.xls`, `.xlsm`, or `.xlsb` — GridGrind handles `.xlsx` only
- Your work is truly one-off and hand-writing JSON adds more friction than it saves
- You need interactive formula recalculation during editing (GridGrind evaluates on request)

## One request, one result

A single JSON request describes every step: create a sheet, write cells, assert workbook state,
read facts back, and save. GridGrind executes the steps in order and writes the file only when
every step succeeds. If an assertion fails or any step errors, nothing is saved.

The request below creates a sheet, writes a cell, asserts the workbook has no malformed hyperlinks,
and reads the written cell back:

```json
{
  "protocolVersion": "V1",
  "source": { "type": "NEW" },
  "persistence": { "type": "SAVE_AS", "path": "lots.xlsx" },
  "execution": {
    "mode": { "type": "FULL_XSSF" },
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

Write, assert, read — one plan. The file is saved only if every step passes.

## Documentation

- [Full docs index](docs/INDEX.md) — every reference file organized by topic
- [First run guide](docs/QUICK_START.md) — first successful run, Docker or JAR
- [Snippets](docs/QUICK_REFERENCE.md) — copy-paste request patterns
- [Java authoring](docs/JAVA_AUTHORING.md) — build requests from Java instead of JSON
- [Operations reference](docs/OPERATIONS.md) — every field and operation
- [Examples](examples/) — ready-to-run request files

## Legal

GridGrind is MIT-licensed. Its executable JAR bundles third-party components under Apache 2.0,
BSD 2-Clause, BSD 3-Clause, and EDL 1.0 licenses. See [NOTICE](NOTICE) for the complete
attribution list and [PATENTS.md](PATENTS.md) for patent considerations.

[LICENSE](LICENSE) | [NOTICE](NOTICE) | [PATENTS.md](PATENTS.md) | [LICENSE-APACHE-2.0](LICENSE-APACHE-2.0) | [LICENSE-BSD-2-CLAUSE](LICENSE-BSD-2-CLAUSE) | [LICENSE-BSD-3-CLAUSE](LICENSE-BSD-3-CLAUSE) | [LICENSE-EDL-1.0](LICENSE-EDL-1.0)
