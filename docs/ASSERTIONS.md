---
afad: "4.0"
version: "0.62.0"
domain: ASSERTIONS
updated: "2026-05-01"
route:
  keywords: [gridgrind, assertions, expect-cell-value, expect-display-value, expect-analysis-max-severity]
  questions: ["what assertions does gridgrind support", "how do assertions work in gridgrind", "how do i verify workbook facts in gridgrind"]
---

# Assertion Reference

**Purpose**: Detailed reference for ordered GridGrind assertion steps and their response/failure
model.
**Landing page**: [ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[WORKBOOK_AND_CELL_INSPECTIONS.md](./WORKBOOK_AND_CELL_INSPECTIONS.md),
[DRAWING_AND_STRUCTURED_INSPECTIONS.md](./DRAWING_AND_STRUCTURED_INSPECTIONS.md), and
[ANALYSIS_QUERIES.md](./ANALYSIS_QUERIES.md)

## Assertions

Assertions are ordered, explicit verification steps. They run in the same `steps[]` list as
mutations and inspections, and they can appear anywhere the plan needs them. Every assertion must
include a caller-defined `stepId`.

Successful responses echo passed assertion steps back through the ordered `assertions[]` array:

```json
{
  "status": "SUCCESS",
  "protocolVersion": "V1",
  "journal": {
    "planId": "assert-budget",
    "level": "NORMAL"
  },
  "persistence": {
    "type": "NONE"
  },
  "warnings": [],
  "assertions": [
    {
      "stepId": "assert-title",
      "assertionType": "EXPECT_CELL_VALUE"
    }
  ],
  "inspections": []
}
```

Failed assertions stop the workflow with `ASSERTION_FAILED` and attach a structured
`problem.assertionFailure` payload describing the failed assertion plus the observed factual read
results that caused the mismatch.

Entity-presence assertions are selector-count assertions, not strict read lookups. If an exact
named-range, chart, table, or pivot-table selector matches nothing, the assertion observes zero
entities and then passes or fails from that count; GridGrind does not surface selector-specific
`*_NOT_FOUND` problems for these assertion families.

Assertion families:

| Assertion `type` | Valid target families | Purpose |
|:-----------------|:----------------------|:--------|
| `EXPECT_NAMED_RANGE_PRESENT` | `NamedRangeSelector` | Require at least one matching named range; selector misses count as zero observed entities. |
| `EXPECT_NAMED_RANGE_ABSENT` | `NamedRangeSelector` | Require zero matching named ranges; selector misses count as zero observed entities. |
| `EXPECT_TABLE_PRESENT` | `TableSelector` | Require at least one matching table; selector misses count as zero observed entities. |
| `EXPECT_TABLE_ABSENT` | `TableSelector` | Require zero matching tables; selector misses count as zero observed entities. |
| `EXPECT_PIVOT_TABLE_PRESENT` | `PivotTableSelector` | Require at least one matching pivot table; selector misses count as zero observed entities. |
| `EXPECT_PIVOT_TABLE_ABSENT` | `PivotTableSelector` | Require zero matching pivot tables; selector misses count as zero observed entities. |
| `EXPECT_CHART_PRESENT` | `ChartSelector` | Require at least one matching chart; selector misses count as zero observed entities. |
| `EXPECT_CHART_ABSENT` | `ChartSelector` | Require zero matching charts; selector misses count as zero observed entities. |
| `EXPECT_CELL_VALUE` | `CellSelector` | Require exact effective cell values. |
| `EXPECT_DISPLAY_VALUE` | `CellSelector` | Require exact formatted display strings. |
| `EXPECT_FORMULA_TEXT` | `CellSelector` | Require exact stored formula text. |
| `EXPECT_CELL_STYLE` | `CellSelector` | Require exact cell-style reports. |
| `EXPECT_WORKBOOK_PROTECTION` | `WorkbookSelector` | Require exact workbook-protection facts. |
| `EXPECT_SHEET_STRUCTURE` | `SheetSelector.ByName` | Require exact one-sheet structural summary facts. |
| `EXPECT_NAMED_RANGE_FACTS` | `NamedRangeSelector` | Require exact named-range reports. |
| `EXPECT_TABLE_FACTS` | `TableSelector` | Require exact table reports. |
| `EXPECT_PIVOT_TABLE_FACTS` | `PivotTableSelector` | Require exact pivot-table reports. |
| `EXPECT_CHART_FACTS` | `ChartSelector` | Require exact chart reports. |
| `EXPECT_ANALYSIS_MAX_SEVERITY` | Matches the supplied analysis `query` target family | Require a maximum severity ceiling for one analysis query. |
| `EXPECT_ANALYSIS_FINDING_PRESENT` | Matches the supplied analysis `query` target family | Require at least one matching finding from one analysis query. |
| `EXPECT_ANALYSIS_FINDING_ABSENT` | Matches the supplied analysis `query` target family | Require no matching findings from one analysis query. |
| `ALL_OF` | Any selector family shared by all nested assertions | Require every nested assertion to pass against the same target. |
| `ANY_OF` | Any selector family shared by all nested assertions | Require at least one nested assertion to pass against the same target. |
| `NOT` | Same selector family as the nested assertion | Invert one nested assertion. |

Common assertion step shapes:

```json
{
  "stepId": "assert-title",
  "target": {
    "type": "CELL_BY_ADDRESS",
    "sheetName": "Budget",
    "address": "A1"
  },
  "assertion": {
    "type": "EXPECT_CELL_VALUE",
    "expectedValue": {
      "type": "TEXT",
      "text": "Quarterly Budget"
    }
  }
}
```

```json
{
  "stepId": "assert-formula-health",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Budget"
  },
  "assertion": {
    "type": "EXPECT_ANALYSIS_MAX_SEVERITY",
    "query": {
      "type": "ANALYZE_FORMULA_HEALTH"
    },
    "maximumSeverity": "INFO"
  }
}
```

```json
{
  "stepId": "assert-workbook-links",
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

```json
{
  "stepId": "assert-total-state",
  "target": {
    "type": "CELL_BY_ADDRESS",
    "sheetName": "Budget",
    "address": "B2"
  },
  "assertion": {
    "type": "ALL_OF",
    "assertions": [
      {
        "type": "EXPECT_CELL_VALUE",
        "expectedValue": {
          "type": "NUMBER",
          "number": 1200
        }
      },
      {
        "type": "NOT",
        "assertion": {
          "type": "EXPECT_DISPLAY_VALUE",
          "displayValue": "0"
        }
      }
    ]
  }
}
```

For a runnable mutate-then-verify plan, see
[`examples/assertion-request.json`](../examples/assertion-request.json).

---
