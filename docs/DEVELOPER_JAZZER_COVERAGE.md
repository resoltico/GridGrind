---
afad: "3.5"
version: "0.58.0"
domain: DEVELOPER_JAZZER_COVERAGE
updated: "2026-04-23"
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
| `protocol-request` | `GridGrindJson.readRequest(byte[])` | raw JSON parsing and request validation | Yes | Yes | 45 |
| `protocol-workflow` | `DefaultGridGrindRequestExecutor.execute(...)` | ordered request workflows through the production contract-plus-executor layer | Yes | Yes | 11 |
| `engine-command-sequence` | `WorkbookCommandExecutor.apply(...)` | ordered workbook-command execution in the engine layer | Yes | Yes | 9 |
| `xlsx-roundtrip` | `ExcelWorkbook.save(...)` plus POI reopen | `.xlsx` persistence and reopen invariants after bounded command sequences | Yes | Yes | 26 |
| `regression` | four isolated per-harness regression tasks over all committed promoted inputs | replay of the committed custom seed floor | N/A | Yes | 91 total across harnesses |

---

## Harness Matrix

### `protocol-request`

Surface:
- raw request bytes
- JSON decoding
- request-model validation
- top-level `formulaEnvironment` payloads for external workbooks, missing-workbook policy, and
  template-backed UDF toolpacks
- top-level `execution.mode` payloads for `EVENT_READ` and `STREAMING_WRITE`, plus
  `execution.journal.level`
- ordered `steps` payloads, selectors, and inspection-step/result correlation IDs
- style payloads including typed `fontHeight`, structured color, fill, gradient, and border input
  shapes
- hyperlink, comment, named-range, data-validation, table, and autofilter payload shapes

What it asserts:
- no unexpected crash during parsing
- valid payloads produce a non-null `WorkbookPlan`
- invalid JSON and invalid request shapes are classified as expected-invalid outcomes
- promoted public examples continue to parse with their real defaulted-field contract, such as
  `SET_TABLE` requests that omit `showTotalsRow`, the formula-environment request, richer factual
  readback examples, the low-memory large-file example, and the advanced workbook-core mutation
  example

Telemetry signals:
- iteration count
- success vs expected-invalid counts
- error families
- assertion-kind coverage for successfully parsed assertion-bearing request payloads
- read-kind coverage for successfully parsed request payloads
- style-kind coverage for style-bearing requests that parse successfully

What it does not cover:
- workbook execution semantics
- engine behavior
- `.xlsx` round-trips

### `protocol-workflow`

Surface:
- ordered step sequences generated from raw bytes
- `DefaultGridGrindRequestExecutor.execute(...)`
- formula-environment-aware execution setup for external workbook bindings, missing-workbook
  policy, and template-backed UDF toolpacks
- response-shape invariants
- explicit inspection-step execution and ordered inspection-result shaping
- source-type and persistence-type combinations for `.xlsx` workflows
- hyperlink, comment, named-range, data-validation, and formula-lifecycle operations in protocol
  execution paths

What it asserts:
- generated workflows remain constructor-valid or are rejected as generated-invalid inputs
- execution returns a response whose protocol invariants hold
- response persistence/path semantics remain coherent for `NEW`, `EXISTING`, `NONE`, `SAVE_AS`,
  and `OVERWRITE`
- success vs failure response families are recorded

Telemetry signals:
- operation-kind counts
- assertion-kind counts
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
- hyperlink, comment, named-range, data-validation, table, autofilter, and formula-lifecycle
  command execution paths

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
- table, autofilter, and formula-cache persistence boundaries after bounded command sequences

What it asserts:
- successful command sequences can be saved and reopened
- reopened workbook exposes coherent structural metadata
- style-bearing command effects survive reopen for supported formatting-depth fields
- hyperlink and comment metadata remain coherent after reopen when those commands succeed
- persisted OOXML comment refs remain canonical after reopen, including column-edit collision cases
- copied picture and embedded-object OOXML relations remain canonical after reopen, including
  worksheet `objectPr` preview refs and copied-sheet relation-id retargeting cases
- named ranges remain coherent after reopen, including normalized target ordering for reversed
  input ranges
- data-validation state remains readable and normalized after reopen when those commands succeed
- table and autofilter commands remain structurally coherent after reopen when those commands
  succeed or are classified as expected-invalid
- explicit formula-cache clearing removes persisted cached formula values before reopen
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

Committed custom seeds currently in source control. This list is exhaustive and should match the
checked-in `*Inputs` directories exactly.

### `protocol-request` (45)

- `advanced_mutation_request.json`
- `advanced_readback_request.json`
- `budget_request.json`
- `chart_request.json`
- `clear_on_empty_cells.json`
- `conditional_formatting_request.json`
- `data_validation_request.json`
- `delete_last_sheet.json`
- `drawing_media_request.json`
- `duplicate_request_id.json`
- `excel_authoring_essentials_request.json`
- `file_hyperlink_health_request.json`
- `formatting_depth_request.json`
- `formula_environment_request.json`
- `formula_equals_prefix.json`
- `fractional_integer_field.json`
- `get_cells_invalid_address.json`
- `get_cells_out_of_bounds_address.json`
- `get_window_overflow.json`
- `introspection_analysis_request.json`
- `invalid_data_validation_empty_explicit_list.json`
- `invalid_email_no_at_sign.json`
- `invalid_font_height_request.json`
- `invalid_request_shape_missing_request_id.json`
- `invalid_request_shape_null_primitive_boolean.json`
- `invalid_request_shape_null_primitive_int.json`
- `invalid_request_shape_unknown_read_type.json`
- `large_file_modes_request.json`
- `live_workflow_create.json`
- `live_workflow_revise.json`
- `package_security_request.json`
- `pane_and_print_reset_request.json`
- `pivot_request.json`
- `rich_text_request.json`
- `row_column_structure_request.json`
- `schema_empty_sheet.json`
- `schema_formula_cells.json`
- `sheet_management_request.json`
- `sheet_name_too_long.json`
- `source_backed_input_request.json`
- `structural_layout_request.json`
- `table_autofilter_request.json`
- `unknown_field_rejection.json`
- `window_size_limit_exceeded.json`
- `workbook_health_request.json`

### `protocol-workflow` (11)

These are opaque generator-byte cases, so they use neutral case identifiers and rely on replay
metadata for their authoritative decoded semantics.

- `workflow_case_01.bin`
- `workflow_case_02.bin`
- `workflow_case_03.bin`
- `workflow_case_04.bin`
- `workflow_case_05.bin`
- `workflow_case_06.bin`
- `workflow_case_07.bin`
- `workflow_case_08.bin`
- `workflow_case_09.bin`
- `workflow_case_10.bin`
- `workflow_case_11.bin`

### `engine-command-sequence` (9)

- `apply_style_alignment_success.bin`
- `apply_style_formatting_depth_success.bin`
- `create_sheet_only_success.bin`
- `create_sheet_set_range_success.bin`
- `delete_last_visible_sheet_invalid.bin`
- `invalid_cell_address_case.bin`
- `invalid_formula_parser_state_case.bin`
- `overlapping_collapsed_then_expanded_group_columns_expected_invalid.bin`
- `set_hyperlink_clear_comment_invalid_sheet.bin`

### `xlsx-roundtrip` (26)

- `append_row_datetime_style_patch_roundtrip_success.bin`
- `append_row_preserves_styled_blank_row_roundtrip_success.bin`
- `apply_style_alignment_roundtrip_success.bin`
- `apply_style_formatting_depth_roundtrip_success.bin`
- `border_all_none_missing_color_roundtrip_success.bin`
- `chart_formula_reference_cache_roundtrip_success.bin`
- `chart_graphic_frame_relation_lookup_roundtrip_success.bin`
- `chart_title_numeric_cache_roundtrip_success.bin`
- `clear_sheet_protection_unprotected_roundtrip_success.bin`
- `comment_collision_after_delete_columns_roundtrip_success.bin`
- `copy_sheet_embedded_object_relation_retarget_roundtrip_success.bin`
- `copy_sheet_picture_relation_retarget_roundtrip_success.bin`
- `create_sheet_roundtrip_case.bin`
- `create_sheet_set_range_roundtrip_case.bin`
- `embedded_object_copy_sheet_roundtrip_success.bin`
- `embedded_object_empty_package_bytes_roundtrip_success.bin`
- `hyperlink_comment_invalid_row_case.bin`
- `named_range_normalization_roundtrip_success.bin`
- `named_range_shift_overwrite_invalid.bin`
- `partial_collapsed_column_ungroup_roundtrip_success.bin`
- `redundant_noop_ungroup_columns_roundtrip_success.bin`
- `set_hyperlink_replacement_roundtrip_success.bin`
- `set_table_missing_sheet_invalid_roundtrip.bin`
- `single_sheet_roundtrip_case.bin`
- `table_header_rewrite_roundtrip_success.bin`
- `table_header_style_display_roundtrip_success.bin`

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
- promoted request-seed coverage for conditional formatting, drawing media, charts, pivots,
  low-memory modes, and OOXML package security
- deterministic `.xlsx` round-trip preservation checks for drawings, charts, pivots, and
  conditional formatting alongside the existing table, autofilter, validation, and style cases
- readable public example coverage for the advanced readback contract, including workbook
  protection, rich comments, advanced print setup, autofilter criteria or sort state, and table
  metadata
- explicit regression coverage for workbook-view invariants such as rejecting deletion of the
  last visible sheet
- normalized local-file hyperlink path semantics in request seeds and `.xlsx` round-trip
  invariants
- explicit regression coverage for comment-collision repair during structural column edits, plus
  persisted-comment OOXML invariants that reject duplicate reopened comment refs
- explicit regression coverage for copied-picture relation repair, plus persisted drawing-picture
  OOXML invariants that reject saved `a:blip/@r:embed` refs with missing `/xl/media/*` targets
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
- direct selector-sweep seam coverage for the extracted operation-sequence generator/value-factory
  layers outside live native fuzzing

What it does not yet give us:
- exhaustive product-surface fuzzing
- advanced Excel feature-object fuzzing
- CI-enforced fuzz regression

That is intentional. The layer is local, isolated, and explicit by design.
