---
afad: "3.5"
version: "0.45.0"
domain: INVENTORY
updated: "2026-04-13"
route:
  keywords: [gridgrind, apache-poi, poi, xssf, xlsx, capability, inventory, workbook-protection, comments, drawings, charts, pivots, pictures, embedded-objects, print-layout, style, autofilter, table, conditional-formatting, package-security, encryption, signing, streaming]
  questions: ["what xlsx features does gridgrind support", "how does gridgrind compare to apache poi xssf", "does gridgrind support workbook protection", "does gridgrind support rich comments", "does gridgrind support charts", "does gridgrind support pivot tables", "does gridgrind support pictures or embedded objects", "does gridgrind support encrypted xlsx files", "what xlsx capabilities are not exposed in gridgrind"]
---

# Apache POI XSSF `.xlsx` Capability Inventory

**Purpose**: Public capability map between GridGrind's current `.xlsx` contract and Apache POI
`5.5.1` XSSF.

**Scope**:
- `.xlsx` only
- Apache POI XSSF, XSSF event-model, and SXSSF families only when GridGrind exposes them directly
- excludes `.xls`, `.xlsm`, `.xlsb`, and non-spreadsheet POI modules

**Interpretation rules**:
- GridGrind is not a 1:1 clone of the POI Java API.
- This document describes the current public GridGrind contract.
- Where Apache POI itself documents only limited support, GridGrind targets that limited extent,
  not full Microsoft Excel parity.

**Verification snapshot**:
- Apache POI version in this repo: `5.5.1` via `gradle/libs.versions.toml`
- GridGrind workspace checked: `2026-04-13`
- Verification rerun in this workspace:
  - `./gradlew parity`
  - `./gradlew check`
  - `./gradlew --project-dir jazzer check`
  - `jazzer/bin/fuzz-protocol-request -PjazzerMaxDuration=15s --console=plain`
  - `jazzer/bin/fuzz-protocol-workflow -PjazzerMaxDuration=15s --console=plain`
  - `./check.sh`

## Status Legend

| Status | Meaning |
|:-------|:--------|
| `COMPLETE` | Shipped in the current GridGrind `.xlsx` contract and backed by current verification |
| `PARTIAL` | GridGrind exposes an important subset of the family, but some related POI-supported behavior is still outside the public contract |
| `NOT_EXPOSED` | No user-facing GridGrind `.xlsx` surface for the capability |

## Quick Summary

Safe to assume in current GridGrind `.xlsx` workflows:
- workbook create, open, save, overwrite, and workbook-summary reads
- sheet lifecycle, visibility, active or selected state, sheet protection, and workbook protection
- typed cell writes, rich text, formulas, workbook-wide or targeted evaluation, formula-cache
  clearing, recalc-on-open, external workbook bindings, missing-workbook policy control, and
  template-backed UDF toolpacks
- style mutation and style snapshot reads, including theme, indexed, tint, and gradient semantics
- merges, sizing, row or column structure edits, panes, zoom, sheet presentation, print layout,
  and advanced page setup
- comments, hyperlinks, named ranges, data validations, autofilters, tables, conditional
  formatting, workbook analysis, and schema reads
- picture, shape, chart, and embedded-object authoring together with factual drawing and chart
  reads
- limited pivot-table authoring, factual reads, and health analysis to the extent POI XSSF
  documents and supports
- low-memory `EVENT_READ` summary reads and `STREAMING_WRITE` append-oriented authoring within the
  documented request-shape limits
- OOXML encrypted open or save, package-signature inspection, unchanged-signature preservation,
  and explicit re-signing on signed mutations

Currently not exposed:
- threaded comments
- XML-mapped tables and slicers
- drawing-family sheet copy

## Capability Matrix

### Workbook and Sheet Surface

| Capability | POI XSSF baseline | GridGrind status | GridGrind evidence | Notes |
|:-----------|:------------------|:-----------------|:-------------------|:------|
| `.xlsx` workbook create, open, save | Documented usermodel support | `COMPLETE` | `protocol/dto/GridGrindRequest.java`, `engine/excel/ExcelWorkbook.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelWorkbookTest.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | GridGrind hard-restricts the product contract to `.xlsx`. |
| Sheet create, rename, delete, move | Documented guide/API support | `COMPLETE` | `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelWorkbook.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | Core sheet lifecycle is fully surfaced. |
| Sheet state: active sheet, selected sheets, visibility | Documented guide/API support | `COMPLETE` | `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelSheetStateController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelSheetStateControllerTest.java` | This is the shipped sheet-state contract. |
| Sheet protection including optional password hashes | Documented API support | `COMPLETE` | `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelSheetStateController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelSheetStateControllerTest.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | Lock flags and optional password-bearing sheet protection are both surfaced. |
| Workbook protection including workbook or revisions passwords | Documented API support | `COMPLETE` | `protocol/dto/WorkbookProtectionInput.java`, `protocol/read/WorkbookReadOperation.java`, `engine/excel/ExcelSheetStateController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelSheetStateControllerTest.java` | GridGrind can author, clear, and read workbook or revisions lock state plus password-hash presence. |
| Sheet copy | Documented API support | `PARTIAL` | `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelSheetCopyController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelSheetCopyControllerTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java` | GridGrind copies workbook-core structures that POI copies safely, including comments, tables, validations, conditional formatting, local formula-defined names, print layout, and protection metadata. Drawing-family content such as pictures and charts is not copied. |
| Merge and unmerge regions | Documented guide/API support | `COMPLETE` | `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelSheet.java`, `protocol/read/WorkbookReadOperation.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | Mutation and factual reads both ship. |
| Row and column sizing plus auto-size | Documented guide/API support | `COMPLETE` | `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelSheet.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | Explicit widths, explicit heights, and auto-size are surfaced. |
| Row and column structure editing | Documented guide/API support | `COMPLETE` | `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelRowColumnStructureController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelRowColumnStructureControllerTest.java`, `jazzer/src/main/java/dev/erst/gridgrind/jazzer/support/XlsxRoundTripVerifier.java` | Insert, delete, shift, hide, group, and ungroup ship with explicit safety guards where POI itself cannot normalize safely. |
| Panes and zoom | Documented guide/API support | `COMPLETE` | `protocol/dto/PaneInput.java`, `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelSheet.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | Freeze, split, and integer zoom are surfaced. |
| Sheet presentation and sheet-level defaults | Documented XSSF sheet-view and worksheet APIs | `COMPLETE` | `protocol/dto/SheetPresentationInput.java`, `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelSheetPresentationController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelSheetPresentationControllerTest.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | `SET_SHEET_PRESENTATION` and `GET_SHEET_LAYOUT.presentation` surface screen display flags, right-to-left layout, tab color, outline-summary placement, default row or column sizing, and ignored-error suppression. |
| Print layout and advanced page setup | Documented guide/API support plus broader page-setup API surface | `COMPLETE` | `protocol/dto/PrintLayoutInput.java`, `engine/excel/ExcelPrintLayoutController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelPrintLayoutControllerTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java` | Print area, titles, headers, footers, orientation, scaling, margins, print-gridline output, centering, paper size, draft, black-and-white, copies, first-page number, and row or column breaks all ship. |

### Cell Content, Styles, and Formulas

| Capability | POI XSSF baseline | GridGrind status | GridGrind evidence | Notes |
|:-----------|:------------------|:-----------------|:-------------------|:------|
| Typed single-cell and range writes plus append row | Documented guide/API support | `COMPLETE` | `protocol/dto/CellInput.java`, `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelSheet.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | Shipped cell families are blank, text, rich text, number, boolean, formula, date, and date-time. |
| Rich text inside a cell | Documented guide/API support | `COMPLETE` | `protocol/dto/RichTextRunInput.java`, `engine/excel/ExcelRichTextSupport.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelRichTextSupportTest.java` | Ordered runs with optional per-run font overrides ship for cell content. |
| Cell snapshots, windows, schemas, and workbook or sheet summaries | Documented usermodel facts | `COMPLETE` | `protocol/read/WorkbookReadOperation.java`, `engine/excel/ExcelSheet.java`, `protocol/dto/GridGrindResponse.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | This is the core factual read surface for exact workbook state. |
| Style patch and style snapshot reporting | Documented guide/API support | `COMPLETE` | `protocol/dto/CellStyleInput.java`, `engine/excel/WorkbookStyleRegistry.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelCellStyleTest.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | Number formats, alignment, wrapping, borders, fills, font attributes, rotation, indentation, and protection flags ship. |
| Theme, indexed, tint, and gradient style semantics | POI low-level style support | `COMPLETE` | `protocol/dto/ColorInput.java`, `protocol/dto/CellGradientFillInput.java`, `engine/excel/WorkbookStyleRegistry.java`, `engine/src/test/java/dev/erst/gridgrind/excel/WorkbookStyleRegistryTest.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | GridGrind authors and reads structured color bases plus gradient fills without collapsing them to RGB-only approximations. |
| Formula assignment, workbook-wide evaluation, targeted evaluation, cache clearing, and recalc-on-open | Documented formula and evaluation support | `COMPLETE` | `protocol/dto/CellInput.java`, `protocol/dto/FormulaCellTargetInput.java`, `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelWorkbook.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelFormulaEnvironmentTest.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/FormulaEnvironmentRequestExecutorTest.java` | GridGrind supports whole-workbook evaluation, explicit-cell evaluation, explicit persisted-cache clearing, and recalc-on-open. |
| External-workbook formula evaluation setup | Documented in POI formula-evaluation docs | `COMPLETE` | `protocol/dto/FormulaEnvironmentInput.java`, `protocol/dto/FormulaExternalWorkbookInput.java`, `engine/excel/ExcelFormulaRuntime.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/FormulaEnvironmentRequestExecutorTest.java` | Request-scoped external workbook bindings drive cross-workbook evaluation and refreshed cached results. |
| Missing-workbook policy control | Documented in POI formula-evaluation docs | `COMPLETE` | `protocol/dto/FormulaEnvironmentInput.java`, `protocol/dto/FormulaMissingWorkbookPolicy.java`, `engine/excel/ExcelFormulaRuntime.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/FormulaEnvironmentRequestExecutorTest.java` | GridGrind exposes strict failure versus cached-value fallback for unresolved external references. |
| User-defined formula functions | Documented in POI UDF guide | `COMPLETE` | `protocol/dto/FormulaUdfToolpackInput.java`, `protocol/dto/FormulaUdfFunctionInput.java`, `engine/excel/ExcelFormulaRuntime.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/FormulaEnvironmentRequestExecutorTest.java` | Template-backed UDF toolpacks provide request-scoped user-defined function execution without exposing arbitrary code loading. |

### Comments, Hyperlinks, Names, Tables, Filters, and Validation

| Capability | POI XSSF baseline | GridGrind status | GridGrind evidence | Notes |
|:-----------|:------------------|:-----------------|:-------------------|:------|
| Cell comments with author, visibility, rich-text runs, and anchors | Documented guide/API support | `COMPLETE` | `protocol/dto/CommentInput.java`, `engine/excel/ExcelCommentSupport.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelCommentTest.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelRichTextSupportTest.java` | GridGrind authors plain text plus optional runs and optional anchor bounds, and factual reads surface those structures back. |
| Threaded comments | Specialized XSSF support beyond classic comments | `NOT_EXPOSED` | No protocol DTO or engine surface for threaded comments found | Classic comment support ships; threaded comments do not. |
| Hyperlink target authoring and reads | Documented guide/API support | `COMPLETE` | `protocol/dto/HyperlinkTarget.java`, `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelSheet.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelHyperlinkTest.java` | URL, email, file, and document targets ship. |
| Named-range authoring for explicit or formula-defined targets | Documented guide/API support | `COMPLETE` | `protocol/dto/NamedRangeTarget.java`, `engine/excel/ExcelWorkbook.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelWorkbookTest.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java` | GridGrind writes workbook- and sheet-scoped names backed either by explicit ranges or by formulas. |
| Named-range factual reads and health analysis | POI usermodel facts | `COMPLETE` | `protocol/read/WorkbookReadOperation.java`, `engine/excel/WorkbookAnalyzer.java`, `protocol/exec/WorkbookReadResultConverter.java`, `engine/src/test/java/dev/erst/gridgrind/excel/WorkbookAnalyzerTest.java` | Reads and analysis cover both range-backed and formula-backed names. |
| Data-validation authoring, factual reads, and health analysis | Documented guide/API support | `COMPLETE` | `protocol/dto/DataValidationInput.java`, `engine/excel/ExcelDataValidationController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelDataValidationControllerTest.java` | Supported families are explicit list, formula list, whole number, decimal, date, time, text length, and custom formula. Prompt, error-alert, and date/time edge states are normalized and surfaced truthfully. |
| Fine-grained data-validation diagnostics including malformed or empty-explicit-list states | POI can load richer edge states than a basic CRUD surface | `COMPLETE` | `protocol/dto/DataValidationEntryReport.java`, `engine/excel/WorkbookAnalyzer.java`, `jazzer/src/main/java/dev/erst/gridgrind/jazzer/support/WorkbookInvariantChecks.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DataValidationEntryReportTest.java` | Distinct malformed-rule and empty-explicit-list states are preserved and surfaced instead of being collapsed into generic unsupported buckets. |
| Autofilter range, criteria, sort-state, and health analysis | Documented guide/API support plus broader XSSF surface | `COMPLETE` | `protocol/dto/AutofilterFilterColumnInput.java`, `protocol/dto/AutofilterSortStateInput.java`, `engine/excel/ExcelAutofilterController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelAutofilterControllerTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java` | GridGrind authors and reads persisted filter criteria and sort state for sheet-level autofilters. Sparse criteria, omitted sort-state metadata, and defaulted Top10 details stay normalized instead of collapsing into misleading unsupported output. |
| Table lifecycle, style, totals-row state, advanced metadata, and health analysis | Documented XSSF table API support | `COMPLETE` | `protocol/dto/TableInput.java`, `protocol/dto/TableColumnInput.java`, `engine/excel/ExcelTableController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelTableControllerTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java` | GridGrind covers table metadata beyond the basic range and style contract, including autofilter state, comments, published or insert-row flags, cell-style metadata, and per-column totals or calculated-formula metadata. |
| XML-mapped tables and slicers | Broader XSSF table ecosystem | `NOT_EXPOSED` | No protocol or engine surface for XML mapping or slicers found | Table authoring is complete for the current public contract, but XML mapping and slicers are not exposed. |

### Drawings, Charts, and Embedded Objects

| Capability | POI XSSF baseline | GridGrind status | GridGrind evidence | Notes |
|:-----------|:------------------|:-----------------|:-------------------|:------|
| Drawing-object factual reads and truthful anchor inspection | XSSF drawing usermodel and OOXML-backed read surface | `COMPLETE` | `protocol/dto/DrawingObjectReport.java`, `protocol/read/WorkbookReadOperation.java`, `engine/excel/ExcelDrawingController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelDrawingControllerTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java` | GridGrind reads pictures, simple shapes, connectors, groups, graphic frames, and embedded objects with factual `TWO_CELL`, `ONE_CELL`, or `ABSOLUTE` anchors. Preview-image relations, embedded-object metadata, and blank-title drawing facts are normalized truthfully on reopen instead of disappearing behind generic null handling. |
| Picture authoring, anchor mutation, delete, and binary payload extraction | Documented quick-guide and XSSF drawing APIs | `COMPLETE` | `protocol/dto/PictureInput.java`, `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelDrawingController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelDrawingControllerTest.java`, `jazzer/src/test/java/dev/erst/gridgrind/jazzer/support/XlsxRoundTripVerifierTest.java` | GridGrind can create, replace, move, delete, read, and extract named pictures with explicit authored two-cell anchors. |
| Simple-shape and connector authoring plus authored anchor replacement | Documented XSSF shape APIs | `COMPLETE` | `protocol/dto/ShapeInput.java`, `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelDrawingController.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java`, `jazzer/src/main/java/dev/erst/gridgrind/jazzer/support/WorkbookInvariantChecks.java` | The shipped authored shape boundary is intentionally honest: `SIMPLE_SHAPE` and `CONNECTOR` only, both backed by explicit two-cell anchors. |
| Embedded-object authoring, preservation, and extracted payload reads | Documented quick-guide embedded-object APIs to the extent POI supports them | `COMPLETE` | `protocol/dto/EmbeddedObjectInput.java`, `protocol/dto/DrawingObjectPayloadReport.java`, `engine/excel/ExcelDrawingController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelDrawingControllerTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityOracle.java` | GridGrind creates named embedded objects with authoritative package bytes and preview images, preserves them across unrelated edits, and extracts truthful user-visible payload bytes on readback. |
| Comment and drawing coexistence without package corruption | Interplay between XSSF comments, VML, and spreadsheet drawings | `COMPLETE` | `engine/excel/ExcelDrawingController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelDrawingControllerTest.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelSheetCopyControllerTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java` | Comment mutation no longer risks deleting or corrupting unrelated drawing relationships, even on reopened workbooks that already contain POI-authored comment or drawing parts. |
| Chart discovery, factual reads, and chart-backed drawing inventory | POI documents limited chart support and chart readback through XDDF/XSSF APIs | `COMPLETE` | `protocol/dto/ChartReport.java`, `protocol/read/WorkbookReadOperation.java`, `protocol/dto/DrawingObjectReport.java`, `engine/excel/ExcelDrawingController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelDrawingControllerTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java` | `GET_CHARTS` returns supported simple `BAR`, `LINE`, and `PIE` charts authoritatively and surfaces unsupported plot families as explicit `UNSUPPORTED` entries. Blank static titles, cached formula titles, and adjacent drawing-frame state are all preserved through truthful rereads. |
| Simple chart authoring, targeted mutation, named-range-backed series binding, and anchor integrity | POI documents limited simple-chart creation and editing | `COMPLETE` | `protocol/dto/ChartInput.java`, `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelChartDefinition.java`, `engine/excel/ExcelDrawingController.java`, `examples/chart-request.json`, `jazzer/src/test/java/dev/erst/gridgrind/jazzer/support/XlsxRoundTripVerifierTest.java` | GridGrind authors and mutates supported simple charts by name, binds series to contiguous ranges or defined names, keeps chart relations or explicit anchor updates intact across save or reopen, and rejects invalid chart payloads without leaving partial chart frames behind. |
| Unsupported chart preservation during unrelated workbook edits | POI loads more chart families than it can authoritatively mutate | `COMPLETE` | `engine/src/test/java/dev/erst/gridgrind/excel/ExcelDrawingControllerTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityOracle.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java`, `protocol/src/parityTest/resources/dev/erst/gridgrind/protocol/parity/xssf-parity-ledger.json` | Unsupported loaded charts reopen as explicit `UNSUPPORTED` detail and survive unrelated edits non-lossily. |

### Pivot Tables, Conditional Formatting, and Analysis

| Capability | POI XSSF baseline | GridGrind status | GridGrind evidence | Notes |
|:-----------|:------------------|:-----------------|:-------------------|:------|
| Pivot-table discovery, factual reads, and pivot-health analysis | POI quick guide and XSSF APIs expose limited pivot discovery and readback support | `COMPLETE` | `protocol/dto/PivotTableReport.java`, `protocol/dto/PivotTableHealthReport.java`, `protocol/read/WorkbookReadOperation.java`, `engine/excel/ExcelPivotTableController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelPivotTableControllerTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java` | `GET_PIVOT_TABLES` returns supported pivots authoritatively and surfaces malformed or unsupported loaded detail as explicit `UNSUPPORTED` entries. `ANALYZE_PIVOT_TABLE_HEALTH` reports missing cache parts, missing workbook cache definitions, broken sources, duplicate or synthetic names, and unsupported persisted detail. |
| Pivot-table authoring and limited mutation for range, named-range, and table sources | POI documents limited pivot creation through XSSF APIs | `COMPLETE` | `protocol/dto/PivotTableInput.java`, `protocol/operation/WorkbookOperation.java`, `engine/excel/ExcelPivotTableDefinition.java`, `engine/excel/ExcelPivotTableController.java`, `examples/pivot-request.json`, `jazzer/src/test/java/dev/erst/gridgrind/jazzer/support/XlsxRoundTripVerifierTest.java` | GridGrind authors workbook-global pivot tables by name from contiguous ranges, existing named ranges, or existing tables, validates disjoint field-role assignment up front, preserves authored cache wiring across save or reopen, rebuilds workbook cache registration on replace or delete, and exposes explicit delete semantics guarded by expected sheet ownership. |
| Pivot-table preservation during unrelated workbook edits | POI-created pivot caches and relations must survive adjacent workbook-core edits non-lossily | `COMPLETE` | `engine/src/test/java/dev/erst/gridgrind/excel/ExcelPivotTableControllerTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityOracle.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java`, `protocol/src/parityTest/resources/dev/erst/gridgrind/protocol/parity/xssf-parity-ledger.json` | Supported pivot tables survive unrelated workbook edits and reopen flows without losing cache definitions, workbook cache references, or persisted source metadata. |
| Conditional-format authoring for formula, cell-value, color-scale, data-bar, icon-set, and top-10 families | Documented guide/API support | `COMPLETE` | `protocol/dto/ConditionalFormattingRuleInput.java`, `engine/excel/ExcelConditionalFormattingController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelConditionalFormattingControllerTest.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/ConditionalFormattingInputTest.java` | GridGrind exposes supported conditional-format families honestly and rejects unsupported authored semantics instead of pretending to preserve them. |
| Conditional-format factual reads and health analysis for loaded rules | POI can load richer rule families | `COMPLETE` | `protocol/dto/ConditionalFormattingRuleReport.java`, `protocol/read/WorkbookReadOperation.java`, `engine/excel/ExcelConditionalFormattingController.java`, `engine/src/test/java/dev/erst/gridgrind/excel/WorkbookAnalyzerTest.java` | Reads and health analysis report formula, cell-value, color-scale, data-bar, icon-set, top-10, and unsupported-rule structures already present in a workbook. Healthy multi-formula rules stay distinct from actually broken formulas, so no-finding branches remain truthful. |
| Formula surface, sheet schema, named-range surface, and aggregate workbook findings | GridGrind-derived analysis over POI facts | `COMPLETE` | `protocol/read/WorkbookReadOperation.java`, `engine/excel/WorkbookAnalyzer.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/DefaultGridGrindRequestExecutorTest.java`, `engine/src/test/java/dev/erst/gridgrind/excel/WorkbookAnalyzerTest.java` | These are value-added GridGrind reads built on top of the factual workbook surface. |

### Large-File Modes and Package Security

| Capability | POI XSSF baseline | GridGrind status | GridGrind evidence | Notes |
|:-----------|:------------------|:-----------------|:-------------------|:------|
| Streaming write with `SXSSF` | Documented by POI as a separate large-file mode | `COMPLETE` | `protocol/dto/ExecutionModeInput.java`, `protocol/exec/DefaultGridGrindRequestExecutor.java`, `engine/excel/ExcelStreamingWorkbookWriter.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/ExecutionModeRequestExecutorTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java`, `examples/large-file-modes-request.json` | GridGrind exposes `executionMode.writeMode=STREAMING_WRITE` for low-memory append-oriented authoring on `NEW` workbooks. The contract is intentionally narrow: only `ENSURE_SHEET`, `APPEND_ROW`, and `FORCE_FORMULA_RECALC_ON_OPEN`. |
| SAX or event-style `.xlsx` reads | Documented by POI for large-file streaming reads | `COMPLETE` | `protocol/dto/ExecutionModeInput.java`, `protocol/exec/DefaultGridGrindRequestExecutor.java`, `engine/excel/ExcelEventWorkbookReader.java`, `engine/src/test/java/dev/erst/gridgrind/excel/ExcelEventWorkbookReaderTest.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/ExecutionModeRequestExecutorTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java` | GridGrind exposes `executionMode.readMode=EVENT_READ` for low-memory `GET_WORKBOOK_SUMMARY` and `GET_SHEET_SUMMARY` requests. The validator makes the summary-only boundary explicit and never silently falls back to the full workbook reader, even when workbook-view, dimension, row-ref, or column-max metadata is sparse. |
| OOXML encryption, password-protected package open or save, and XML signing | Official POI encryption guide documents these families | `COMPLETE` | `protocol/dto/OoxmlOpenSecurityInput.java`, `protocol/dto/OoxmlPersistenceSecurityInput.java`, `protocol/read/WorkbookReadOperation.java`, `engine/excel/ExcelOoxmlPackageSecuritySupport.java`, `protocol/src/test/java/dev/erst/gridgrind/protocol/exec/OoxmlSecurityRequestExecutorTest.java`, `protocol/src/parityTest/java/dev/erst/gridgrind/protocol/parity/XlsxParityProbeRegistry.java`, `examples/package-security-inspect-request.json` | GridGrind opens encrypted `.xlsx` packages through `source.security.password`, emits encrypted or signed `.xlsx` outputs through `persistence.security`, returns factual `GET_PACKAGE_SECURITY` reports, preserves unchanged signatures, and requires explicit re-signing before persisting signed mutations. Pass-through eligibility, signing failure translation, and persisted package-state rereads are regression-covered. |

## Important Notes

- This document is intentionally `.xlsx`-only. Mixed-format rows for `.xls`, `.xlsm`, and `.xlsb`
  stay out because they obscure the actual public contract.
- `COMPLETE` means complete for the current shipped GridGrind contract, not that every POI API
  reachable from XSSF has been productized.
- Status notes call out narrower boundaries where the surrounding POI family is broader than the
  public GridGrind surface.

## Primary Apache POI Sources Used

- [Spreadsheet Overview](https://poi.apache.org/components/spreadsheet/)
- [Busy Developers' Guide to HSSF and XSSF Features](https://poi.apache.org/components/spreadsheet/quick-guide.html)
- [HSSF and XSSF Limitations](https://poi.apache.org/components/spreadsheet/limitations.html)
- [Formula Evaluation](https://poi.apache.org/components/spreadsheet/eval.html)
- [User Defined Functions](https://poi.apache.org/components/spreadsheet/user-defined-functions.html)
- [Encryption Support](https://poi.apache.org/encryption.html)
- [History of Changes](https://poi.apache.org/changes.html)
- [XSSFSheet API](https://poi.apache.org/apidocs/dev/org/apache/poi/xssf/usermodel/XSSFSheet.html)
- [XSSFTable API](https://poi.apache.org/apidocs/dev/org/apache/poi/xssf/usermodel/XSSFTable.html)
