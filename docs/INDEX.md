---
afad: "4.0"
version: "0.65.0"
domain: INDEX
updated: "2026-05-16"
route:
  keywords: [gridgrind, index, docs, documentation, reference, map]
  questions: ["where is the gridgrind documentation index", "what docs does gridgrind have", "how is the gridgrind documentation organized"]
---

# Documentation Index

Complete map of every file in `docs/`. Files are grouped by audience and topic.

---

## Start Here

| File | What it covers |
|:-----|:---------------|
| [QUICK_START.md](./QUICK_START.md) | First successful run — Docker or JAR, new workbook, read result back |
| [QUICK_REFERENCE.md](./QUICK_REFERENCE.md) | Copy-paste request fragments: cells, styles, assertions, inspections, charts |
| [OPERATIONS.md](./OPERATIONS.md) | Stable public index of every mutation action, assertion, and inspection query; links to each detail reference |
| [EXAMPLES.md](./EXAMPLES.md) | Runnable example files in `examples/`, path-rooting rules, and refresh flow |

---

## Request And Execution

| File | What it covers |
|:-----|:---------------|
| [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md) | Request envelope fields (`source`, `persistence`, `execution`, `formulaEnvironment`), source-backed inputs, doctor requests, coordinate systems, cell value shapes, response journal |
| [LIMITATIONS.md](./LIMITATIONS.md) | Hard ceilings and mode restrictions (LIM-001 through LIM-037) |
| [ERRORS.md](./ERRORS.md) | Problem model, error codes, and recovery guidance |

---

## Mutation Reference

### Landing Pages

Two-tier structure: landing pages give the overview; detail pages own the full field reference.

| File | What it covers |
|:-----|:---------------|
| [WORKBOOK_AND_LAYOUT_MUTATIONS.md](./WORKBOOK_AND_LAYOUT_MUTATIONS.md) | Landing page: workbook lifecycle, sheet management, layout and structure |
| [CELL_AND_DRAWING_MUTATIONS.md](./CELL_AND_DRAWING_MUTATIONS.md) | Landing page: cell values, hyperlinks, comments, pictures, shapes, charts |
| [STRUCTURED_FEATURE_MUTATIONS.md](./STRUCTURED_FEATURE_MUTATIONS.md) | Landing page: styles, validations, conditional formatting, tables, pivot tables, autofilters, named ranges |

### Detail Pages

| File | What it covers |
|:-----|:---------------|
| [WORKBOOK_AND_SHEET_MUTATIONS.md](./WORKBOOK_AND_SHEET_MUTATIONS.md) | `ENSURE_SHEET`, `RENAME_SHEET`, `DELETE_SHEET`, `MOVE_SHEET`, `COPY_SHEET`, sheet visibility, protection, active/selected sheets, custom XML import, workbook protection |
| [LAYOUT_AND_STRUCTURE_MUTATIONS.md](./LAYOUT_AND_STRUCTURE_MUTATIONS.md) | Merges, row/column insert/delete/shift, visibility, grouping, panes, zoom, presentation, print layout |
| [CELL_VALUE_MUTATIONS.md](./CELL_VALUE_MUTATIONS.md) | `SET_CELL`, `SET_RANGE`, `SET_ARRAY_FORMULA`, `CLEAR_ARRAY_FORMULA`, `CLEAR_RANGE` |
| [LINK_AND_COMMENT_MUTATIONS.md](./LINK_AND_COMMENT_MUTATIONS.md) | `SET_HYPERLINK`, `CLEAR_HYPERLINK`, `SET_COMMENT`, `CLEAR_COMMENT` |
| [DRAWING_MUTATIONS.md](./DRAWING_MUTATIONS.md) | `SET_PICTURE`, `SET_SHAPE`, `SET_EMBEDDED_OBJECT`, `SET_CHART`, `SET_SIGNATURE_LINE`, `SET_DRAWING_OBJECT_ANCHOR`, `DELETE_DRAWING_OBJECT` |
| [STYLE_AND_VALIDATION_MUTATIONS.md](./STYLE_AND_VALIDATION_MUTATIONS.md) | `APPLY_STYLE`, `SET_DATA_VALIDATION`, `CLEAR_DATA_VALIDATIONS`, `SET_CONDITIONAL_FORMATTING`, `CLEAR_CONDITIONAL_FORMATTING` |
| [STRUCTURED_DATA_MUTATIONS.md](./STRUCTURED_DATA_MUTATIONS.md) | `SET_AUTOFILTER`, `CLEAR_AUTOFILTER`, `SET_TABLE`, `DELETE_TABLE`, `SET_PIVOT_TABLE`, `DELETE_PIVOT_TABLE`, `APPEND_ROW`, `AUTO_SIZE_COLUMNS`, `SET_NAMED_RANGE`, `DELETE_NAMED_RANGE` |

---

## Assertion And Inspection Reference

| File | What it covers |
|:-----|:---------------|
| [ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md) | Landing page: all assertion and inspection families with cross-links to detail pages |
| [ASSERTIONS.md](./ASSERTIONS.md) | Every assertion type (`EXPECT_CELL_VALUE`, `EXPECT_ANALYSIS_MAX_SEVERITY`, `ALL_OF`, `NOT`, etc.), response shape, failure payload |
| [WORKBOOK_AND_CELL_INSPECTIONS.md](./WORKBOOK_AND_CELL_INSPECTIONS.md) | Workbook, sheet, cell, range, and package-security factual reads |
| [DRAWING_AND_STRUCTURED_INSPECTIONS.md](./DRAWING_AND_STRUCTURED_INSPECTIONS.md) | Picture, shape, embedded-object, chart, signature-line, table, pivot, autofilter, and named-range factual reads |
| [ANALYSIS_QUERIES.md](./ANALYSIS_QUERIES.md) | Workbook-health and analysis query payloads (`ANALYZE_WORKBOOK_FINDINGS`, `ANALYZE_FORMULA_HEALTH`, `ANALYZE_NAMED_RANGE_HEALTH`, etc.) |

---

## Java Authoring

| File | What it covers |
|:-----|:---------------|
| [JAVA_AUTHORING.md](./JAVA_AUTHORING.md) | Fluent Java plan building, selector helpers, source-backed inputs, JSON emission, and optional explicit executor handoff |

---

## Developer Documentation

| File | What it covers |
|:-----|:---------------|
| [DEVELOPER.md](./DEVELOPER.md) | Architecture, module map, build commands, GitHub workflows, quality gates, JaCoCo |
| [DEVELOPER_GRADLE.md](./DEVELOPER_GRADLE.md) | Gradle 9 project structure, task catalog, dependency rules |
| [DEVELOPER_JAVA.md](./DEVELOPER_JAVA.md) | Java 26 conventions, sealed types, records, Jackson 3.x usage |
| [DEVELOPER_DOCKER.md](./DEVELOPER_DOCKER.md) | Docker image build, run, and testing |
| [DEVELOPER_DEVCONTAINER.md](./DEVELOPER_DEVCONTAINER.md) | Dev container setup and usage |
| [DEVELOPER_JAZZER.md](./DEVELOPER_JAZZER.md) | Jazzer fuzz-testing overview and status |
| [DEVELOPER_JAZZER_OPERATIONS.md](./DEVELOPER_JAZZER_OPERATIONS.md) | Running Jazzer fuzz jobs |
| [DEVELOPER_JAZZER_COVERAGE.md](./DEVELOPER_JAZZER_COVERAGE.md) | Jazzer corpus and coverage guidance |
| [DEVELOPER_CONTRACT_REPLACEMENT_ADR.md](./DEVELOPER_CONTRACT_REPLACEMENT_ADR.md) | Architecture decision record: contract module replacement approach |
| [RELEASE_PROTOCOL.md](./RELEASE_PROTOCOL.md) | Release checklist and publishing steps |

---

## Capability Inventory

| File | What it covers |
|:-----|:---------------|
| [POI_EXCEL_CAPABILITY_INVENTORY.md](./POI_EXCEL_CAPABILITY_INVENTORY.md) | Apache POI / XSSF capability audit: what is and is not supported at the POI layer |
