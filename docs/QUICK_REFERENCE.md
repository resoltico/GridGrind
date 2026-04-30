---
afad: "4.0"
version: "0.62.0"
domain: QUICK_REFERENCE
updated: "2026-05-01"
route:
  keywords: [gridgrind, quick-reference, snippets, request, execution, examples, formula, workbook-health, chart, signature-line]
  questions: ["what is the quickest way to write a gridgrind request", "how do I generate a built-in gridgrind example", "what are the most common gridgrind request snippets", "where is the detailed gridgrind reference"]
---

# Quick Reference

Fast-start snippets and reminders for the shipped GridGrind `.xlsx` contract. Use this file as a
cheat sheet, then jump to the detailed references when you need the full field list.

## Artifact Discovery

```bash
gridgrind help
gridgrind --print-request-template --response request.json
gridgrind --print-protocol-catalog --response protocol-catalog.json
gridgrind --print-protocol-catalog --search validation --response validation-search.json
gridgrind --print-protocol-catalog --operation inspectionQueryTypes:GET_SHEET_LAYOUT
gridgrind --print-example BUDGET --response budget-request.json
gridgrind --print-example SHEET_MAINTENANCE --response sheet-maintenance.json
gridgrind --print-example WORKBOOK_HEALTH --response workbook-health.json
gridgrind --print-example ASSERTION --response assertion.json
gridgrind --print-request-template | gridgrind --doctor-request
gridgrind --doctor-request --request request.json --response doctor-report.json
```

`gridgrind help` is the explicit alias for `--help`. `--doctor-request` validates request shape,
resolves source-backed inputs, and preflights existing workbook-source access without mutating a
workbook. `--response <path>` works across execution, doctoring, and discovery commands, so the
primary output can be captured to a file instead of stdout.

`--search` is the fast discovery path when you only know part of an id or summary. Use
`--operation <group>:<id>` once you want one exact machine-readable entry.

## Smallest Valid Request

```json
{
  "protocolVersion": "V1",
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "NONE"
  },
  "execution": {
    "mode": {
      "readMode": "FULL_XSSF",
      "writeMode": "FULL_XSSF"
    },
    "journal": {
      "level": "NORMAL"
    },
    "calculation": {
      "strategy": {
        "type": "DO_NOT_CALCULATE"
      },
      "markRecalculateOnOpen": false
    }
  },
  "formulaEnvironment": {
    "externalWorkbooks": [],
    "missingWorkbookPolicy": "ERROR",
    "udfToolpacks": []
  },
  "steps": []
}
```

Every non-empty step needs a caller-defined `stepId`. Step kind is inferred from exactly one of
`action`, `assertion`, or `query`; do not send `step.type`.
`gridgrind --print-request-template` emits this same canonical scaffold. The step snippets below
are request fragments, not standalone full requests, unless the section explicitly shows the full
top-level envelope.

## Common Source And Persistence Shapes

Create a new workbook:

```json
{ "source": { "type": "NEW" } }
```

Open an existing workbook:

```json
{ "source": { "type": "EXISTING", "path": "budget.xlsx" } }
```

Open an encrypted workbook:

```json
{
  "source": {
    "type": "EXISTING",
    "path": "secured.xlsx",
    "security": {
      "password": "GridGrind-2026"
    }
  }
}
```

Save to a new path:

```json
{ "persistence": { "type": "SAVE_AS", "path": "out/report.xlsx" } }
```

Run without saving:

```json
{ "persistence": { "type": "NONE" } }
```

## Source-Backed Authored Inputs

Common text and binary sources:

```json
{ "type": "INLINE", "text": "Quarterly note" }
{ "type": "UTF8_FILE", "path": "authored-inputs/note.txt" }
{ "type": "STANDARD_INPUT" }
{ "type": "INLINE_BASE64", "base64Data": "SGVsbG8=" }
{ "type": "FILE", "path": "authored-inputs/payload.bin" }
{ "type": "STANDARD_INPUT" }
```

The request JSON transport is capped at `16 MiB`. Large authored text or binary content belongs in
`UTF8_FILE`, `FILE`, or `STANDARD_INPUT` sources instead of inline JSON strings.

When the CLI reads a request via `--request <path>`, relative request-owned paths such as
`source.path`, `persistence.path`, `UTF8_FILE` / `FILE`, external workbook bindings, and signing
material resolve from that request file's directory. `--request` and `--response` themselves
still resolve from the shell working directory.

## Execution, Formula, And Mode Rules

Common execution block:

```json
{
  "execution": {
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

Targeted evaluation:

```json
{
  "execution": {
    "calculation": {
      "strategy": {
        "type": "EVALUATE_TARGETS",
        "cells": [
          {
            "sheetName": "Budget",
            "address": "D2"
          },
          {
            "sheetName": "Budget",
            "address": "E2"
          }
        ]
      }
    }
  }
}
```

Rules to remember:
- `execution.mode.readMode=EVENT_READ` supports `GET_WORKBOOK_SUMMARY` and `GET_SHEET_SUMMARY`
  only.
- `execution.mode.writeMode=STREAMING_WRITE` requires `source.type=NEW` and supports only
  `ENSURE_SHEET` plus `APPEND_ROW`.
- `markRecalculateOnOpen` persists Excel's workbook-open recalculation flag.
- `EVALUATE_TARGETS` addresses must point at existing formula cells. Missing physical cells can
  surface `CELL_NOT_FOUND`; existing non-formula cells are rejected as `INVALID_REQUEST`.
- Scalar `FORMULA` cell payloads reject array-formula braces such as `{=SUM(A1:A2*B1:B2)}`.
  Use `SET_ARRAY_FORMULA` for contiguous array groups.
- `LAMBDA` and `LET` are currently rejected as `INVALID_FORMULA` because Apache POI cannot parse
  them yet.

## Common Step Snippets

Ensure a sheet exists:

```json
{
  "stepId": "ensure-budget",
  "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
  "action": { "type": "ENSURE_SHEET" }
}
```

Write one cell:

```json
{
  "stepId": "set-title",
  "target": {
    "type": "CELL_BY_ADDRESS",
    "sheetName": "Budget",
    "address": "A1"
  },
  "action": {
    "type": "SET_CELL",
    "value": { "type": "TEXT", "source": { "type": "INLINE", "text": "Quarterly Budget" } }
  }
}
```

Write a row range:

```json
{
  "stepId": "set-header-row",
  "target": {
    "type": "RANGE_BY_RANGE",
    "sheetName": "Budget",
    "range": "A2:C2"
  },
  "action": {
    "type": "SET_RANGE",
    "rows": [
      [
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Team"
          }
        },
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Plan"
          }
        },
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Actual"
          }
        }
      ]
    ]
  }
}
```

Apply a style patch:

```json
{
  "stepId": "style-header",
  "target": { "type": "RANGE_BY_RANGE", "sheetName": "Budget", "range": "A2:C2" },
  "action": {
    "type": "APPLY_STYLE",
    "style": { "bold": true, "fill": { "pattern": "SOLID", "foregroundColor": "#D9EAF7" } }
  }
}
```

Advanced drawing mutations live in [DRAWING_MUTATIONS.md](./DRAWING_MUTATIONS.md):
`SET_PICTURE`, `SET_SHAPE`, `SET_EMBEDDED_OBJECT`, `SET_SIGNATURE_LINE`, `SET_CHART`,
`SET_DRAWING_OBJECT_ANCHOR`, and `DELETE_DRAWING_OBJECT`.

Supported authored plot families are `AREA`, `AREA_3D`, `BAR`, `BAR_3D`, `DOUGHNUT`, `LINE`,
`LINE_3D`, `PIE`, `PIE_3D`, `RADAR`, `SCATTER`, `SURFACE`, and `SURFACE_3D`.

For full runnable chart and signature-line requests, start from
[../examples/chart-request.json](../examples/chart-request.json) and
[../examples/signature-line-request.json](../examples/signature-line-request.json).

## Common Assertion And Inspection Snippets

Exact cell-value assertion:

```json
{
  "stepId": "assert-total",
  "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "D2" },
  "assertion": {
    "type": "EXPECT_CELL_VALUE",
    "expectedValue": { "type": "NUMBER", "number": 1200 }
  }
}
```

Workbook-health assertion:

```json
{
  "stepId": "assert-workbook-clean",
  "target": {
    "type": "WORKBOOK_CURRENT"
  },
  "assertion": {
    "type": "EXPECT_ANALYSIS_FINDING_ABSENT",
    "query": {
      "type": "ANALYZE_WORKBOOK_FINDINGS"
    },
    "code": "HYPERLINK_MALFORMED_TARGET"
  }
}
```

Read a small cell set:

```json
{
  "stepId": "cells",
  "target": {
    "type": "CELL_BY_ADDRESSES",
    "sheetName": "Budget",
    "addresses": [
      "A1",
      "D2"
    ]
  },
  "query": {
    "type": "GET_CELLS"
  }
}
```

Read sheet layout:

```json
{
  "stepId": "layout",
  "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
  "query": { "type": "GET_SHEET_LAYOUT" }
}
```

Run a no-save workbook-health pass by starting from the smallest valid request above, switching
`source` to `EXISTING`, keeping `persistence.type` as `NONE`, and adding one
`ANALYZE_WORKBOOK_FINDINGS` step against `WORKBOOK_CURRENT`. The full runnable shape lives in
[REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md) and
[../examples/workbook-health-request.json](../examples/workbook-health-request.json).

## Response Anchors Worth Remembering

- `GET_FORMULA_SURFACE`: `analysis.totalFormulaCellCount`
- `ANALYZE_NAMED_RANGE_HEALTH`: `analysis.checkedNamedRangeCount`, `analysis.summary`, and
  `analysis.findings`
- `ANALYZE_WORKBOOK_FINDINGS`: `analysis.summary` and `analysis.findings`
- Every response, success or failure: top-level structured `journal`

## Detailed References

- [JAVA_AUTHORING.md](./JAVA_AUTHORING.md)
- [OPERATIONS.md](./OPERATIONS.md)
- [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md)
- [WORKBOOK_AND_SHEET_MUTATIONS.md](./WORKBOOK_AND_SHEET_MUTATIONS.md)
- [LAYOUT_AND_STRUCTURE_MUTATIONS.md](./LAYOUT_AND_STRUCTURE_MUTATIONS.md)
- [CELL_VALUE_MUTATIONS.md](./CELL_VALUE_MUTATIONS.md)
- [LINK_AND_COMMENT_MUTATIONS.md](./LINK_AND_COMMENT_MUTATIONS.md)
- [DRAWING_MUTATIONS.md](./DRAWING_MUTATIONS.md)
- [STYLE_AND_VALIDATION_MUTATIONS.md](./STYLE_AND_VALIDATION_MUTATIONS.md)
- [STRUCTURED_DATA_MUTATIONS.md](./STRUCTURED_DATA_MUTATIONS.md)
- [ASSERTIONS.md](./ASSERTIONS.md)
- [WORKBOOK_AND_CELL_INSPECTIONS.md](./WORKBOOK_AND_CELL_INSPECTIONS.md)
- [DRAWING_AND_STRUCTURED_INSPECTIONS.md](./DRAWING_AND_STRUCTURED_INSPECTIONS.md)
- [ANALYSIS_QUERIES.md](./ANALYSIS_QUERIES.md)
- [ERRORS.md](./ERRORS.md)
- [LIMITATIONS.md](./LIMITATIONS.md)
