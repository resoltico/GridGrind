---
afad: "3.4"
version: "0.30.0"
domain: DEVELOPER_JAZZER_COVERAGE
updated: "2026-04-07"
route:
  keywords: [gridgrind, jazzer, fuzz, coverage, matrix, harnesses, regression inputs, promoted inputs, gaps]
  questions: ["what does jazzer cover in gridgrind", "which harnesses exist", "what are the promoted jazzer inputs", "what gaps remain in jazzer coverage", "what does each jazzer target assert"]
---

# Jazzer Coverage Inventory

**Purpose**: Current-stock inventory of what the GridGrind Jazzer layer covers, what promoted
regression inputs exist, and what remains outside the current fuzzing surface.
**Architecture reference**: [DEVELOPER_JAZZER.md](./DEVELOPER_JAZZER.md)
**Operator reference**: [DEVELOPER_JAZZER_OPERATIONS.md](./DEVELOPER_JAZZER_OPERATIONS.md)

---

## Coverage Summary

| Target | Entry Point | Concern | Replay Support | Telemetry | Promoted Inputs |
|:-------|:------------|:--------|:---------------|:----------|:----------------|
| `protocol-request` | `GridGrindJson.readRequest(byte[])` | raw JSON parsing and request validation | Yes | Yes | 33 |
| `protocol-workflow` | `DefaultGridGrindRequestExecutor.execute(...)` | ordered request workflows through the production protocol/service layer | Yes | Yes | 11 |
| `engine-command-sequence` | `WorkbookCommandExecutor.apply(...)` | ordered workbook-command execution in the engine layer | Yes | Yes | 8 |
| `xlsx-roundtrip` | `ExcelWorkbook.save(...)` plus POI reopen | `.xlsx` persistence and reopen invariants after bounded command sequences | Yes | Yes | 15 |
| `regression` | four isolated per-harness regression tasks over all committed promoted inputs | replay of the committed custom seed floor | N/A | Yes | 67 total across harnesses |

---

## Harness Matrix

### `protocol-request`

Surface:
- raw request bytes
- JSON decoding
- request-model validation
- ordered `reads` payloads, selectors, and request/result correlation IDs
- style payloads including typed `fontHeight`, fill, color, and border input shapes
- hyperlink, comment, named-range, data-validation, table, and autofilter payload shapes

What it asserts:
- no unexpected crash during parsing
- valid payloads produce a non-null `GridGrindRequest`
- invalid JSON and invalid request shapes are classified as expected-invalid outcomes

Telemetry signals:
- iteration count
- success vs expected-invalid counts
- error families
- read-kind coverage for successfully parsed request payloads
- style-kind coverage for style-bearing requests that parse successfully

What it does not cover:
- workbook execution semantics
- engine behavior
- `.xlsx` round-trips

### `protocol-workflow`

Surface:
- ordered operation sequences generated from raw bytes
- `DefaultGridGrindRequestExecutor.execute(...)`
- response-shape invariants
- explicit read execution and ordered read-result shaping
- source-type and persistence-type combinations for `.xlsx` workflows
- hyperlink, comment, named-range, and data-validation operations in protocol execution paths

What it asserts:
- generated workflows remain constructor-valid or are rejected as generated-invalid inputs
- execution returns a response whose protocol invariants hold
- response persistence/path semantics remain coherent for `NEW`, `EXISTING`, `NONE`, `SAVE_AS`,
  and `OVERWRITE`
- success vs failure response families are recorded

Telemetry signals:
- operation-kind counts
- style-kind counts
- source-kind counts
- persistence-kind counts
- read-kind counts
- success vs generated-invalid counts
- response-family counts
- unexpected failure families

What it does not cover:
- direct engine-only paths bypassing `DefaultGridGrindRequestExecutor`
- `.xlsx` reopen after save

### `engine-command-sequence`

Surface:
- ordered `WorkbookCommand` sequences generated from raw bytes
- direct `WorkbookCommandExecutor.apply(...)`
- workbook structural invariants
- hyperlink, comment, named-range, data-validation, table, and autofilter command execution paths

What it asserts:
- workbook shape remains coherent after successful execution
- invalid command sequences fail within expected validation families
- style-bearing commands remain constructor-valid or fail within expected validation families

Telemetry signals:
- command-kind counts
- style-kind counts
- success vs expected-invalid counts
- error families

What it does not cover:
- protocol JSON parsing
- protocol response shaping
- `.xlsx` reopen after save

### `xlsx-roundtrip`

Surface:
- ordered `WorkbookCommand` sequences
- save to `.xlsx`
- reopen with Apache POI `XSSFWorkbook`
- table and autofilter persistence boundaries after bounded command sequences

What it asserts:
- successful command sequences can be saved and reopened
- reopened workbook exposes coherent structural metadata
- style-bearing command effects survive reopen for supported formatting-depth fields
- hyperlink and comment metadata remain coherent after reopen when those commands succeed
- named ranges remain coherent after reopen, including normalized target ordering for reversed
  input ranges
- data-validation state remains readable and normalized after reopen when those commands succeed
- table and autofilter commands remain structurally coherent after reopen when those commands
  succeed or are classified as expected-invalid
- table header rewrites, header clears, and header-range style patches keep persisted table-column
  metadata converged with the visible header cells after reopen

Telemetry signals:
- command-kind counts
- style-kind counts
- success vs expected-invalid counts
- error families

What it does not cover:
- charts
- encrypted or macro-bearing workbooks

---

## Deterministic Support Tests

The nested Jazzer build now also includes deterministic support tests under:

```text
jazzer/src/test/java/
```

Current deterministic support scope:
- `ReplayGridGrindFuzzDataTest`: scalar replay semantics for the pure-Java deterministic replay
  cursor used by structured binary-harness replay
- `JazzerReportSupportTest`: latest-summary parsing, including active-fuzz corpus-size parsing
- `JazzerTextRendererTest`: summary rendering semantics for active findings vs replay-clean
  artifacts
- `JazzerRegressionRunnerTest`: committed-seed replay launch behavior, including non-JSON harness
  replay without Jazzer's native replay provider
- `JazzerReplaySupportTest`: replay-time classification, stable replay expectations, and artifact
  classification for committed raw findings
- `WorkbookInvariantChecksTest`: protocol-workflow and response-shape invariants outside the fuzz
  loop
- `XlsxRoundTripVerifierTest`: targeted style-aware and authoring-metadata round-trip verifier
  behavior outside the fuzz loop, including expectation derivation from the real pre-save workbook
  state
- `PromotionMetadataTest`: replay-expectation verification for every promoted metadata entry

These tests are not fuzz harnesses. They protect the Jazzer infrastructure itself.

---

## Current Promoted Input Inventory

Committed custom seeds currently in source control:

| Harness | Input | Meaning |
|:--------|:------|:--------|
| `protocol-request` | `sheet_management_request.json` | readable valid request seed taken from the public sheet-management example |
| `protocol-request` | `budget_request.json` | readable budget workflow seed with range writes, style, formulas, and explicit workbook/cell/window/schema reads |
| `protocol-request` | `data_validation_request.json` | readable validation workflow seed covering validation authoring, partial clearing, factual validation reads, and validation-health analysis |
| `protocol-request` | `table_autofilter_request.json` | readable table-and-autofilter workflow seed covering sheet filters, table authoring, factual reads, and both health-analysis families |
| `protocol-request` | `excel_authoring_essentials_request.json` | readable authoring seed covering URL and FILE hyperlinks, comments, named ranges, and explicit metadata plus named-range reads |
| `protocol-request` | `file_hyperlink_health_request.json` | readable hyperlink-analysis seed covering FILE path inputs, `file:` URI normalization, hyperlink metadata reads, and hyperlink-health analysis |
| `protocol-request` | `live_workflow_create.json` | readable multi-sheet finance workflow with append-row and formula authoring |
| `protocol-request` | `live_workflow_revise.json` | readable existing-workbook revision seed with overwrite persistence |
| `protocol-request` | `row_column_structure_request.json` | readable B3 structural-edit seed covering row and column insertion, deletion, shifting, visibility, grouping, and `GET_SHEET_LAYOUT` reads |
| `protocol-request` | `structural_layout_request.json` | readable structural-layout seed covering merge, sizing, pane state, zoom, and print layout |
| `protocol-request` | `formatting_depth_request.json` | readable formatting-depth seed covering typed `fontHeight`, fill, color, and border patches |
| `protocol-request` | `invalid_data_validation_empty_explicit_list.json` | readable expected-invalid seed covering empty explicit-list validation rejection |
| `protocol-request` | `invalid_font_height_request.json` | readable expected-invalid seed covering typed `fontHeight` validation |
| `protocol-request` | `invalid_request_shape_missing_request_id.json` | readable expected-invalid seed covering missing required read fields and `INVALID_REQUEST_SHAPE` replay classification |
| `protocol-request` | `invalid_request_shape_unknown_read_type.json` | readable expected-invalid seed covering unknown read discriminators and `INVALID_REQUEST_SHAPE` replay classification |
| `protocol-workflow` | `set_cell_failure_case.bin` | structured workflow seed that replays to a protocol `FAILURE` response with one `SET_CELL` operation |
| `protocol-workflow` | `ensure_sheet_set_range_success.bin` | structured workflow seed that replays to a protocol `SUCCESS` response with `ENSURE_SHEET` plus `SET_RANGE` |
| `protocol-workflow` | `apply_style_success.bin` | structured workflow seed that replays to a protocol `SUCCESS` response dominated by `APPLY_STYLE` |
| `protocol-workflow` | `append_row_failure.bin` | structured workflow seed that replays to a protocol `FAILURE` response dominated by `APPEND_ROW` |
| `protocol-workflow` | `auto_size_failure.bin` | structured workflow seed that replays to a protocol `FAILURE` response dominated by `AUTO_SIZE_COLUMNS` |
| `protocol-workflow` | `existing_overwrite_apply_style_success.bin` | structured workflow seed that replays through `EXISTING` plus `OVERWRITE` with style application |
| `protocol-workflow` | `existing_overwrite_hyperlink_failure.bin` | structured workflow seed that replays through `EXISTING` plus `OVERWRITE` with hyperlink operations in a protocol failure path |
| `protocol-workflow` | `save_as_apply_style_failure.bin` | structured workflow seed that exercises `SAVE_AS` plus style-bearing protocol failure handling |
| `protocol-workflow` | `save_as_named_range_failure.bin` | structured workflow seed that exercises `SAVE_AS` plus repeated named-range operations in a protocol failure path |
| `protocol-workflow` | `set_comment_clear_range_failure.bin` | structured workflow seed that exercises comment mutation together with later range clearing in a protocol failure path |
| `protocol-workflow` | `existing_overwrite_set_autofilter_failure.bin` | structured workflow seed that replays through `EXISTING` plus `OVERWRITE` with autofilter operations in a protocol failure path |
| `engine-command-sequence` | `invalid_cell_address_case.bin` | direct engine seed that replays to an expected `InvalidCellAddressException` |
| `engine-command-sequence` | `create_sheet_only_success.bin` | direct engine seed that replays to successful repeated `CREATE_SHEET` commands |
| `engine-command-sequence` | `create_sheet_set_range_success.bin` | direct engine seed that replays to successful `CREATE_SHEET` plus `SET_RANGE` commands |
| `engine-command-sequence` | `apply_style_alignment_success.bin` | direct engine seed that replays to successful alignment-only style application |
| `engine-command-sequence` | `apply_style_formatting_depth_success.bin` | direct engine seed that replays to successful formatting-depth style application |
| `engine-command-sequence` | `invalid_formula_parser_state_case.bin` | direct engine seed that replays to the expected-invalid malformed-formula parser-state case fixed from a former Jazzer finding |
| `engine-command-sequence` | `set_hyperlink_clear_comment_invalid_sheet.bin` | direct engine seed that replays to an expected-invalid mixed hyperlink/comment command case against a missing sheet |
| `engine-command-sequence` | `delete_last_visible_sheet_invalid.bin` | direct engine seed promoted from a live B1 finding to assert that deleting the last visible sheet is classified as expected-invalid instead of crashing view-state normalization |
| `xlsx-roundtrip` | `create_sheet_roundtrip_case.bin` | successful round-trip seed dominated by repeated `CREATE_SHEET` commands |
| `xlsx-roundtrip` | `single_sheet_roundtrip_case.bin` | minimal successful round-trip seed that creates one sheet and persists cleanly |
| `xlsx-roundtrip` | `create_sheet_set_range_roundtrip_case.bin` | successful round-trip seed that persists `CREATE_SHEET` plus `SET_RANGE` commands |
| `xlsx-roundtrip` | `apply_style_alignment_roundtrip_success.bin` | successful round-trip seed that preserves alignment-only style state after reopen |
| `xlsx-roundtrip` | `apply_style_formatting_depth_roundtrip_success.bin` | successful round-trip seed that preserves formatting-depth style state after reopen |
| `xlsx-roundtrip` | `append_row_preserves_styled_blank_row_roundtrip_success.bin` | successful round-trip seed promoted from a live finding to assert style preservation when `APPEND_ROW` reuses styled blank rows |
| `xlsx-roundtrip` | `append_row_datetime_style_patch_roundtrip_success.bin` | successful round-trip seed promoted from a live finding to assert that date-time append writes relayer their required number format onto styled blank rows |
| `xlsx-roundtrip` | `table_header_rewrite_roundtrip_success.bin` | successful round-trip seed promoted from a live finding to assert that table header rewrites keep persisted table metadata synchronized through reopen |
| `xlsx-roundtrip` | `table_header_style_display_roundtrip_success.bin` | successful round-trip seed promoted from a live finding to assert that header-range style patches keep typed table header metadata synchronized through reopen |
| `xlsx-roundtrip` | `clear_sheet_protection_unprotected_roundtrip_success.bin` | successful round-trip seed promoted from a live B1 finding to assert that clearing protection on an already unprotected sheet is an idempotent no-op instead of a POI-backed crash |
| `xlsx-roundtrip` | `named_range_normalization_roundtrip_success.bin` | successful round-trip seed that preserves named-range state while normalizing reversed target ordering |
| `xlsx-roundtrip` | `named_range_shift_overwrite_invalid.bin` | expected-invalid round-trip seed proving that destructive row shifts against range-backed named ranges are rejected before POI can rewrite them into broken `#REF!` formulas |
| `xlsx-roundtrip` | `hyperlink_comment_invalid_row_case.bin` | expected-invalid round-trip seed that exercises hyperlink and comment commands alongside invalid row mutation input |
| `xlsx-roundtrip` | `set_hyperlink_replacement_roundtrip_success.bin` | successful round-trip seed that preserves the latest hyperlink target after repeated writes to the same cell |
| `xlsx-roundtrip` | `set_table_missing_sheet_invalid_roundtrip.bin` | expected-invalid round-trip seed that exercises table authoring across missing-sheet failure classification |

Matching promotion metadata lives under:

```text
jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata/
```

Those metadata files document:
- source artifact path
- promoted input path
- replay outcome kind
- replay expectation
- replay text artifact path

---

## Present Strengths

The current Jazzer layer is strongest at:
- ordered operation and command interaction coverage
- validation-boundary discovery for sheet names, ranges, addresses, widths, heights, pane
  coordinates, zoom percentages, and print-layout bounds
- a committed custom seed floor that now includes readable public example requests and replay-
  verified binary workflow seeds
- validation-aware request and round-trip coverage for the data-validation authoring family
- validation-aware request and round-trip coverage for the table and autofilter authoring family
- explicit regression coverage for workbook-view invariants such as rejecting deletion of the
  last visible sheet
- normalized local-file hyperlink path semantics in request seeds and `.xlsx` round-trip
  invariants
- style-aware `.xlsx` round-trip verification for formatting depth, authoring metadata, and
  named-range normalization plus data-validation and table/autofilter persistence boundaries
- explicit source/persistence-type telemetry for service-level workflow fuzzing
- deterministic tests that protect the Jazzer reporting and verifier infrastructure itself
- replay-expectation verification that prevents promoted binary-seed semantics from drifting
- regression preservation of real discovered local inputs
- local operator observability after every supported fuzz run

The most important architectural strength is not just the harnesses. It is the full local loop:
- fuzz
- summarize
- inspect
- replay
- promote
- regression

---

## Current Gaps

Still outside the current Jazzer surface:
- CLI transport fuzzing
- file-path and filesystem persistence boundary fuzzing
- conditional-formatting, chart, picture, and pivot coverage
- extremely large-workbook or streaming-strategy fuzzing
- cross-run corpus-health scoring beyond counts and newest-entry inspection
- automatic conversion of promoted inputs into deterministic main-suite tests
- a replay-verified `.xlsx` round-trip success seed that isolates successful hyperlink/comment
  persistence without unrelated command noise
- committed libFuzzer dictionaries

These are real gaps. The current layer is good, but it is not full project fuzzing coverage.

---

## Practical Interpretation

What the current four-harness stack gives us together:
- request decoder hardening
- protocol-layer workflow composition pressure
- engine-layer composition pressure
- persistence/reopen pressure

What it does not yet give us:
- exhaustive product-surface fuzzing
- advanced Excel feature-object fuzzing
- CI-enforced fuzz regression

That is intentional. The layer is local, isolated, and explicit by design.
