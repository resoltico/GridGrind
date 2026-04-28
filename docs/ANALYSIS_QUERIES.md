---
afad: "3.5"
version: "0.60.0"
domain: ANALYSIS_QUERIES
updated: "2026-04-25"
route:
  keywords: [gridgrind, analysis queries, workbook health, formula health, hyperlink health, named-range health]
  questions: ["what analysis queries does gridgrind support", "how do i run workbook health checks in gridgrind", "how do i inspect findings in gridgrind"]
---

# Analysis Query Reference

**Purpose**: Detailed reference for finding-bearing GridGrind analysis queries and workbook-health
surfaces.
**Landing page**: [ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[ASSERTIONS.md](./ASSERTIONS.md), [WORKBOOK_AND_CELL_INSPECTIONS.md](./WORKBOOK_AND_CELL_INSPECTIONS.md),
and [DRAWING_AND_STRUCTURED_INSPECTIONS.md](./DRAWING_AND_STRUCTURED_INSPECTIONS.md)

### ANALYZE_FORMULA_HEALTH

Reports finding-bearing formula health across one or more sheets. This is where volatile
functions, formula-error results, and evaluation failures surface.

```json
{
  "stepId": "formula-health",
  "target": {
    "type": "SHEET_ALL"
  },
  "query": {
    "type": "ANALYZE_FORMULA_HEALTH"
  }
}
```

Response shape: `{ "analysis": { "checkedFormulaCellCount": ..., "summary": ..., "findings":`
`[...] } }`.

### ANALYZE_DATA_VALIDATION_HEALTH

Reports finding-bearing data-validation health across one or more sheets. Findings include
unsupported rules, broken formulas, and overlapping validation coverage.

```json
{
  "stepId": "data-validation-health",
  "target": {
    "type": "SHEET_BY_NAMES",
    "names": [
      "Inventory",
      "Summary"
    ]
  },
  "query": {
    "type": "ANALYZE_DATA_VALIDATION_HEALTH"
  }
}
```

### ANALYZE_CONDITIONAL_FORMATTING_HEALTH

Reports conditional-formatting findings such as broken formulas, unsupported loaded rule
families, empty target ranges, or priority collisions.

```json
{
  "stepId": "conditional-formatting-health",
  "target": {
    "type": "SHEET_BY_NAMES",
    "names": [
      "Inventory",
      "Summary"
    ]
  },
  "query": {
    "type": "ANALYZE_CONDITIONAL_FORMATTING_HEALTH"
  }
}
```

### ANALYZE_AUTOFILTER_HEALTH

Reports autofilter findings such as invalid ranges, blank header rows, or ownership mismatches
between sheet-level filters and tables.

```json
{
  "stepId": "autofilter-health",
  "target": {
    "type": "SHEET_BY_NAMES",
    "names": [
      "Inventory",
      "Summary"
    ]
  },
  "query": {
    "type": "ANALYZE_AUTOFILTER_HEALTH"
  }
}
```

### ANALYZE_TABLE_HEALTH

Reports table findings such as overlaps, broken ranges, blank or duplicate headers, or style
mismatches.

```json
{
  "stepId": "table-health",
  "target": {
    "type": "TABLE_ALL"
  },
  "query": {
    "type": "ANALYZE_TABLE_HEALTH"
  }
}
```

```json
{
  "stepId": "selected-table-health",
  "target": {
    "type": "TABLE_BY_NAMES",
    "names": [
      "InventoryTable",
      "Trips"
    ]
  },
  "query": {
    "type": "ANALYZE_TABLE_HEALTH"
  }
}
```

### ANALYZE_PIVOT_TABLE_HEALTH

Reports pivot-table findings such as missing cache parts, missing workbook-cache definitions,
broken sources, duplicate names, synthetic fallback names, or unsupported persisted detail.

```json
{
  "stepId": "pivot-health",
  "target": {
    "type": "PIVOT_TABLE_ALL"
  },
  "query": {
    "type": "ANALYZE_PIVOT_TABLE_HEALTH"
  }
}
```

```json
{
  "stepId": "selected-pivot-health",
  "target": {
    "type": "PIVOT_TABLE_BY_NAMES",
    "names": [
      "SalesPivot",
      "NamedPivot"
    ]
  },
  "query": {
    "type": "ANALYZE_PIVOT_TABLE_HEALTH"
  }
}
```

### ANALYZE_HYPERLINK_HEALTH

Reports hyperlink findings such as malformed external targets, missing local file targets,
unresolved relative file targets, or broken document destinations.

Relative `FILE` targets are resolved against the workbook's persisted path when `source` or
`persistence` gives the workbook a filesystem location. When the workbook is still unsaved,
relative `FILE` targets are reported as `HYPERLINK_UNRESOLVED_FILE_TARGET`.

```json
{
  "stepId": "hyperlink-health",
  "target": {
    "type": "SHEET_BY_NAMES",
    "names": [
      "Inventory",
      "Summary"
    ]
  },
  "query": {
    "type": "ANALYZE_HYPERLINK_HEALTH"
  }
}
```

### ANALYZE_NAMED_RANGE_HEALTH

Reports named-range findings such as broken references, unresolved targets, and scope shadowing.

```json
{
  "stepId": "named-range-health",
  "target": {
    "type": "NAMED_RANGE_ALL"
  },
  "query": {
    "type": "ANALYZE_NAMED_RANGE_HEALTH"
  }
}
```

Response shape: `{ "analysis": { "checkedNamedRangeCount": ..., "summary": ..., "findings":`
`[...] } }`.

### ANALYZE_WORKBOOK_FINDINGS

Runs every shipped finding-bearing analysis family across the workbook and returns one aggregated
flat finding list. This is the primary workbook-health check and works especially well with
`persistence.type=NONE` when you want a no-save lint pass.

Response shape: `{ "analysis": { "summary": ..., "findings": [...] } }`.

The aggregate currently includes formula, data validation, conditional formatting, autofilter,
table, pivot-table, hyperlink, and named-range findings.

```json
{
  "stepId": "workbook-findings",
  "target": {
    "type": "WORKBOOK_CURRENT"
  },
  "query": {
    "type": "ANALYZE_WORKBOOK_FINDINGS"
  }
}
```

Formula references to same-request sheet names with spaces should use single quotes, for example
`'Budget Review'!B1`. When execution succeeds, GridGrind reports this style of request issue in
the success response `warnings` array.
