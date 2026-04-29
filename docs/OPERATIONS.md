---
afad: "3.5"
version: "0.61.0"
domain: OPERATIONS
updated: "2026-04-25"
route:
  keywords: [gridgrind, operations, assertions, inspections, reference, mutation, query, request, execution, quick-links]
  questions: ["where is the full gridgrind step reference", "what operations does gridgrind support", "what assertions does gridgrind support", "what inspection queries does gridgrind support"]
---

# Operations, Assertions, and Inspection Reference

**Purpose**: Stable public map of the shipped GridGrind `.xlsx` contract.
**Machine-readable discovery**: `gridgrind --print-protocol-catalog`
**Ranked search**: `gridgrind --print-protocol-catalog --search <text>`
**Copy-paste cheat sheet**: [QUICK_REFERENCE.md](./QUICK_REFERENCE.md)
**Hard ceilings and unsupported boundaries**: [LIMITATIONS.md](./LIMITATIONS.md)

This file is intentionally an index, not the full manual. The long-form reference is split into
focused documents so the public surface stays easier to audit and harder to let drift.

## Contract Discovery

```bash
gridgrind --print-request-template --response request.json
gridgrind --print-protocol-catalog --response protocol-catalog.json
gridgrind --print-protocol-catalog --search chart --response chart-search.json
gridgrind --print-protocol-catalog --operation mutationActionTypes:SET_CELL
gridgrind --print-example BUDGET --response budget-request.json
gridgrind --print-example ASSERTION --response assertion-request.json
printf '%s\n' '{"source":{"type":"NEW"},"steps":[]}' | gridgrind --doctor-request
gridgrind --doctor-request --request request.json --response doctor-report.json
```

`--print-protocol-catalog` is the authoritative machine-readable shape inventory. It publishes the
current request model, every mutation action, every assertion type, every inspection query,
required versus optional fields, and the allowed nested selectors or payload groups for polymorphic
fields. Use `--search` when you only know part of the name or summary, then switch to
`--operation <group>:<id>` for the exact entry once you have the stable qualified id. Discovery,
printed example requests, doctor reports, and normal execution responses all omit absent optional
fields, and request payloads must omit absent fields instead of sending explicit JSON `null`
placeholders.
fields instead of publishing explicit JSON `null` placeholders, so the machine-readable surface is
easier for agents and shell tooling to consume directly.
Task discovery is layered on top of that same catalog surface:
`--print-task-catalog --response tasks.json`, `--print-task-plan <id> --response task-plan.json`,
and `--print-goal-plan "<goal>" --response goal-plan.json` now emit
starter scaffolds for dashboards, tabular reports, data-entry flows, pivot reports, custom XML
workflows, workbook maintenance, and drawing/signature workflows.
`--doctor-request` is the fast preflight path for request shape, execution-mode limits,
source-backed input resolution, and existing workbook-source accessibility; it does not mutate a
workbook. `--response <path>` applies across execution, doctoring, and discovery, so primary
outputs can be written directly to files during artifact, shell, or Docker workflows.

## Canonical Terminology

- **Mutation action**: a `steps[]` entry carrying `action`.
- **Assertion step**: a `steps[]` entry carrying `assertion`.
- **Inspection query**: a `steps[]` entry carrying `query`.
- **Source-backed authored input**: request text or binary content loaded from `INLINE`,
  `UTF8_FILE`, `FILE`, or `STANDARD_INPUT` rather than embedded directly in the JSON request body.
- **Execution policy**: the optional request-level `execution` block that controls low-memory modes,
  journaling, and formula handling.

## Reference Map

| Reference | Owns |
|:----------|:-----|
| [JAVA_AUTHORING.md](./JAVA_AUTHORING.md) | Fluent Java plan building, selector helpers, source-backed inputs, JSON emission, and optional explicit executor handoff over the same canonical `WorkbookPlan` |
| [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md) | Request envelope, `source`, `persistence`, source-backed inputs, `formulaEnvironment`, `execution`, response journal, coordinate systems, and core cell-value shapes |
| [WORKBOOK_AND_SHEET_MUTATIONS.md](./WORKBOOK_AND_SHEET_MUTATIONS.md) | workbook lifecycle, sheet create/rename/delete/move/copy, active/selected sheets, protection, and custom XML import |
| [LAYOUT_AND_STRUCTURE_MUTATIONS.md](./LAYOUT_AND_STRUCTURE_MUTATIONS.md) | merges, row and column inserts/deletes/shifts, visibility, grouping, panes, zoom, presentation, and print layout |
| [CELL_VALUE_MUTATIONS.md](./CELL_VALUE_MUTATIONS.md) | `SET_CELL`, `SET_RANGE`, `SET_ARRAY_FORMULA`, `CLEAR_ARRAY_FORMULA`, and `CLEAR_RANGE` |
| [LINK_AND_COMMENT_MUTATIONS.md](./LINK_AND_COMMENT_MUTATIONS.md) | hyperlink and comment authoring |
| [DRAWING_MUTATIONS.md](./DRAWING_MUTATIONS.md) | pictures, shapes, embedded objects, charts, signature lines, anchor updates, and drawing deletes |
| [STYLE_AND_VALIDATION_MUTATIONS.md](./STYLE_AND_VALIDATION_MUTATIONS.md) | styles, data validations, and conditional formatting |
| [STRUCTURED_DATA_MUTATIONS.md](./STRUCTURED_DATA_MUTATIONS.md) | autofilters, tables, pivot tables, `APPEND_ROW`, `AUTO_SIZE_COLUMNS`, `execution.calculation`, and named ranges |
| [ASSERTIONS.md](./ASSERTIONS.md) | all assertion families and common assertion step shapes |
| [WORKBOOK_AND_CELL_INSPECTIONS.md](./WORKBOOK_AND_CELL_INSPECTIONS.md) | workbook, sheet, cell, range, and package-security factual reads |
| [DRAWING_AND_STRUCTURED_INSPECTIONS.md](./DRAWING_AND_STRUCTURED_INSPECTIONS.md) | picture, shape, embedded-object, chart, signature-line, table, pivot, autofilter, and named-range factual reads |
| [ANALYSIS_QUERIES.md](./ANALYSIS_QUERIES.md) | workbook-health and analysis query payloads |

## Mutation Action Groups

Workbook, sheet, and layout actions:
`ENSURE_SHEET`, `RENAME_SHEET`, `DELETE_SHEET`, `MOVE_SHEET`, `COPY_SHEET`,
`SET_ACTIVE_SHEET`, `SET_SELECTED_SHEETS`, `SET_SHEET_VISIBILITY`,
`SET_SHEET_PROTECTION`, `CLEAR_SHEET_PROTECTION`, `SET_WORKBOOK_PROTECTION`,
`CLEAR_WORKBOOK_PROTECTION`, `IMPORT_CUSTOM_XML_MAPPING`, `MERGE_CELLS`,
`UNMERGE_CELLS`, `SET_COLUMN_WIDTH`, `SET_ROW_HEIGHT`, `INSERT_ROWS`, `DELETE_ROWS`,
`SHIFT_ROWS`, `INSERT_COLUMNS`, `DELETE_COLUMNS`, `SHIFT_COLUMNS`, `SET_ROW_VISIBILITY`,
`SET_COLUMN_VISIBILITY`, `GROUP_ROWS`, `UNGROUP_ROWS`, `GROUP_COLUMNS`, `UNGROUP_COLUMNS`,
`SET_SHEET_PANE`, `SET_SHEET_ZOOM`, `SET_SHEET_PRESENTATION`, `SET_PRINT_LAYOUT`, and
`CLEAR_PRINT_LAYOUT`.

Cell and drawing actions:
`SET_CELL`, `SET_RANGE`, `SET_ARRAY_FORMULA`, `CLEAR_ARRAY_FORMULA`, `CLEAR_RANGE`,
`SET_HYPERLINK`, `CLEAR_HYPERLINK`, `SET_COMMENT`, `CLEAR_COMMENT`, `SET_PICTURE`,
`SET_SHAPE`, `SET_EMBEDDED_OBJECT`, `SET_CHART`, `SET_SIGNATURE_LINE`,
`SET_DRAWING_OBJECT_ANCHOR`, and `DELETE_DRAWING_OBJECT`.

Structured feature actions:
`APPLY_STYLE`, `SET_DATA_VALIDATION`, `CLEAR_DATA_VALIDATIONS`,
`SET_CONDITIONAL_FORMATTING`, `CLEAR_CONDITIONAL_FORMATTING`, `SET_AUTOFILTER`,
`CLEAR_AUTOFILTER`, `SET_TABLE`, `DELETE_TABLE`, `SET_PIVOT_TABLE`, `DELETE_PIVOT_TABLE`,
`APPEND_ROW`, `AUTO_SIZE_COLUMNS`, `SET_NAMED_RANGE`, and `DELETE_NAMED_RANGE`.

## Assertion And Inspection Surface

Assertion families include `EXPECT_NAMED_RANGE_PRESENT`, `EXPECT_NAMED_RANGE_ABSENT`,
`EXPECT_TABLE_PRESENT`, `EXPECT_TABLE_ABSENT`, `EXPECT_PIVOT_TABLE_PRESENT`,
`EXPECT_PIVOT_TABLE_ABSENT`, `EXPECT_CHART_PRESENT`, `EXPECT_CHART_ABSENT`,
`EXPECT_CELL_VALUE`, `EXPECT_DISPLAY_VALUE`, `EXPECT_FORMULA_TEXT`, `EXPECT_CELL_STYLE`,
`EXPECT_WORKBOOK_PROTECTION`, `EXPECT_SHEET_STRUCTURE`, `EXPECT_NAMED_RANGE_FACTS`,
`EXPECT_TABLE_FACTS`, `EXPECT_PIVOT_TABLE_FACTS`, `EXPECT_CHART_FACTS`,
`EXPECT_ANALYSIS_MAX_SEVERITY`, `EXPECT_ANALYSIS_FINDING_PRESENT`,
`EXPECT_ANALYSIS_FINDING_ABSENT`, `ALL_OF`, `ANY_OF`, and `NOT`.

Inspection queries cover workbook facts, sheet facts, drawing/chart facts, table/pivot facts,
package security, named ranges, schemas, and every shipped health-analysis family. The detailed
step shapes live in [ASSERTIONS.md](./ASSERTIONS.md),
[WORKBOOK_AND_CELL_INSPECTIONS.md](./WORKBOOK_AND_CELL_INSPECTIONS.md),
[DRAWING_AND_STRUCTURED_INSPECTIONS.md](./DRAWING_AND_STRUCTURED_INSPECTIONS.md), and
[ANALYSIS_QUERIES.md](./ANALYSIS_QUERIES.md).

Common response anchors:
- `GET_FORMULA_SURFACE` returns `analysis.totalFormulaCellCount` plus grouped sheet summaries.
- `ANALYZE_NAMED_RANGE_HEALTH` returns `analysis.checkedNamedRangeCount`, `analysis.summary`, and
  `analysis.findings`.
- `ANALYZE_WORKBOOK_FINDINGS` returns workbook-wide `analysis.summary` plus one flat
  `analysis.findings` list.

## Current Formula And Execution Boundaries

- Scalar `FORMULA` cell values reject request-authored array-formula braces such as `{=...}`.
  Use `SET_ARRAY_FORMULA` for contiguous array groups.
- `LAMBDA` and `LET` are currently rejected as `INVALID_FORMULA` because Apache POI cannot parse
  them yet.
- `execution.calculation.strategy` accepts `DO_NOT_CALCULATE`, `EVALUATE_ALL`,
  `EVALUATE_TARGETS`, and `CLEAR_CACHES_ONLY`.
- `execution.calculation.markRecalculateOnOpen` persists Excel's workbook-open recalculation flag
  without a dedicated mutation action.
- `EVALUATE_TARGETS` addresses must point at existing formula cells. A missing physical cell can
  surface `CELL_NOT_FOUND`; an existing non-formula cell is rejected as `INVALID_REQUEST`.
- `execution.mode.readMode=EVENT_READ` supports `GET_WORKBOOK_SUMMARY` and `GET_SHEET_SUMMARY`
  only.
- `execution.mode.writeMode=STREAMING_WRITE` requires `source.type=NEW`, limits mutations to
  `ENSURE_SHEET` plus `APPEND_ROW`, and allows `markRecalculateOnOpen=true` only with
  `strategy=DO_NOT_CALCULATE`.

## Advanced Drawing And Chart Surface

Drawing authoring and readback are first-class `.xlsx` surfaces. See
[DRAWING_MUTATIONS.md](./DRAWING_MUTATIONS.md) and
[DRAWING_AND_STRUCTURED_INSPECTIONS.md](./DRAWING_AND_STRUCTURED_INSPECTIONS.md) for the full
shapes covering `SIGNATURE_LINE`, pictures, shapes, embedded objects, and charts.

Supported authored plot families are `AREA`, `AREA_3D`, `BAR`, `BAR_3D`, `DOUGHNUT`, `LINE`,
`LINE_3D`, `PIE`, `PIE_3D`, `RADAR`, `SCATTER`, `SURFACE`, and `SURFACE_3D`. Multi-plot combo
charts are supported when every plot belongs to one of those families.

## Example Entry Points

- First-run workbook creation: [QUICK_START.md](./QUICK_START.md)
- Example map, path-rooting rules, and refresh flow: [EXAMPLES.md](./EXAMPLES.md)
- Java-first authoring without hand-written JSON: [JAVA_AUTHORING.md](./JAVA_AUTHORING.md) and
  [../examples/java-authoring-workflow.java](../examples/java-authoring-workflow.java)
- Budget walkthrough: [../examples/budget-request.json](../examples/budget-request.json) or
  `gridgrind --print-example BUDGET --response budget-request.json`
- Assertion walkthrough: [../examples/assertion-request.json](../examples/assertion-request.json)
- Workbook-health walkthrough:
  [../examples/workbook-health-request.json](../examples/workbook-health-request.json)
- Workbook-maintenance walkthrough:
  [../examples/sheet-maintenance-request.json](../examples/sheet-maintenance-request.json)
- POI/XSSF capability audit: [POI_EXCEL_CAPABILITY_INVENTORY.md](./POI_EXCEL_CAPABILITY_INVENTORY.md)
- Problem model and recovery: [ERRORS.md](./ERRORS.md)
