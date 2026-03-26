---
afad: "3.4"
version: "0.6.0"
domain: DEVELOPER_JAZZER_COVERAGE
updated: "2026-03-26"
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
| `protocol-request` | `GridGrindJson.readRequest(byte[])` | raw JSON parsing and request validation | Yes | Yes | 7 |
| `protocol-workflow` | `GridGrindService.execute(...)` | ordered request workflows through the production protocol/service layer | Yes | Yes | 7 |
| `engine-command-sequence` | `WorkbookCommandExecutor.apply(...)` | ordered workbook-command execution in the engine layer | Yes | Yes | 6 |
| `xlsx-roundtrip` | `ExcelWorkbook.save(...)` plus POI reopen | `.xlsx` persistence and reopen invariants after bounded command sequences | Yes | Yes | 5 |
| `regression` | all committed `*Inputs` resources | replay of the committed custom seed floor | N/A | Yes | 25 total across harnesses |

---

## Harness Matrix

### `protocol-request`

Surface:
- raw request bytes
- JSON decoding
- request-model validation
- style payloads including typed `fontHeight`, fill, color, and border input shapes

What it asserts:
- no unexpected crash during parsing
- valid payloads produce a non-null `GridGrindRequest`
- invalid JSON and invalid request shapes are classified as expected-invalid outcomes

Telemetry signals:
- iteration count
- success vs expected-invalid counts
- error families
- style-kind coverage for style-bearing requests that parse successfully

What it does not cover:
- workbook execution semantics
- engine behavior
- `.xlsx` round-trips

### `protocol-workflow`

Surface:
- ordered operation sequences generated from raw bytes
- `GridGrindService.execute(...)`
- response-shape invariants
- source-mode and persistence-mode combinations for `.xlsx` workflows

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
- success vs generated-invalid counts
- response-family counts
- unexpected failure families

What it does not cover:
- direct engine-only paths bypassing `GridGrindService`
- `.xlsx` reopen after save

### `engine-command-sequence`

Surface:
- ordered `WorkbookCommand` sequences generated from raw bytes
- direct `WorkbookCommandExecutor.apply(...)`
- workbook structural invariants

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

What it asserts:
- successful command sequences can be saved and reopened
- reopened workbook exposes coherent structural metadata
- style-bearing command effects survive reopen for supported formatting-depth fields

Telemetry signals:
- command-kind counts
- style-kind counts
- success vs expected-invalid counts
- error families

What it does not cover:
- charts
- tables
- data validation
- encrypted or macro-bearing workbooks

---

## Deterministic Support Tests

The nested Jazzer build now also includes deterministic support tests under:

```text
jazzer/src/test/java/
```

Current deterministic support scope:
- `JazzerReportSupportTest`: latest-summary parsing, including active-fuzz corpus-size parsing
- `JazzerTextRendererTest`: summary rendering semantics for active findings vs replay-clean
  artifacts
- `XlsxRoundTripVerifierTest`: targeted style-aware round-trip verifier behavior outside the fuzz
  loop

These tests are not fuzz harnesses. They protect the Jazzer infrastructure itself.

---

## Current Promoted Input Inventory

Committed custom seeds currently in source control:

| Harness | Input | Meaning |
|:--------|:------|:--------|
| `protocol-request` | `sheet_management_request.json` | readable valid request seed taken from the public sheet-management example |
| `protocol-request` | `budget_request.json` | readable budget workflow seed with range writes, style, formulas, and analysis |
| `protocol-request` | `live_workflow_create.json` | readable multi-sheet finance workflow with append-row and formula authoring |
| `protocol-request` | `live_workflow_revise.json` | readable existing-workbook revision seed with overwrite persistence |
| `protocol-request` | `structural_layout_request.json` | readable structural-layout seed covering merge, sizing, and freeze panes |
| `protocol-request` | `formatting_depth_request.json` | readable formatting-depth seed covering typed `fontHeight`, fill, color, and border patches |
| `protocol-request` | `invalid_font_height_request.json` | readable expected-invalid seed covering typed `fontHeight` validation |
| `protocol-workflow` | `set_cell_failure_case.bin` | structured workflow seed that replays to a protocol `FAILURE` response with one `SET_CELL` operation |
| `protocol-workflow` | `ensure_sheet_set_range_success.bin` | structured workflow seed that replays to a protocol `SUCCESS` response with `ENSURE_SHEET` plus `SET_RANGE` |
| `protocol-workflow` | `apply_style_success.bin` | structured workflow seed that replays to a protocol `SUCCESS` response dominated by `APPLY_STYLE` |
| `protocol-workflow` | `append_row_failure.bin` | structured workflow seed that replays to a protocol `FAILURE` response dominated by `APPEND_ROW` |
| `protocol-workflow` | `auto_size_failure.bin` | structured workflow seed that replays to a protocol `FAILURE` response dominated by `AUTO_SIZE_COLUMNS` |
| `protocol-workflow` | `existing_overwrite_apply_style_success.bin` | structured workflow seed that replays through `EXISTING` plus `OVERWRITE` with style application |
| `protocol-workflow` | `save_as_apply_style_failure.bin` | structured workflow seed that exercises `SAVE_AS` plus style-bearing protocol failure handling |
| `engine-command-sequence` | `invalid_cell_address_case.bin` | direct engine seed that replays to an expected `InvalidCellAddressException` |
| `engine-command-sequence` | `create_sheet_only_success.bin` | direct engine seed that replays to successful repeated `CREATE_SHEET` commands |
| `engine-command-sequence` | `create_sheet_set_range_success.bin` | direct engine seed that replays to successful `CREATE_SHEET` plus `SET_RANGE` commands |
| `engine-command-sequence` | `apply_style_alignment_success.bin` | direct engine seed that replays to successful alignment-only style application |
| `engine-command-sequence` | `apply_style_formatting_depth_success.bin` | direct engine seed that replays to successful formatting-depth style application |
| `engine-command-sequence` | `invalid_formula_parser_state_case.bin` | direct engine seed that replays to the expected-invalid malformed-formula parser-state case fixed from a former Jazzer finding |
| `xlsx-roundtrip` | `create_sheet_roundtrip_case.bin` | successful round-trip seed dominated by repeated `CREATE_SHEET` commands |
| `xlsx-roundtrip` | `single_sheet_roundtrip_case.bin` | minimal successful round-trip seed that creates one sheet and persists cleanly |
| `xlsx-roundtrip` | `create_sheet_set_range_roundtrip_case.bin` | successful round-trip seed that persists `CREATE_SHEET` plus `SET_RANGE` commands |
| `xlsx-roundtrip` | `apply_style_alignment_roundtrip_success.bin` | successful round-trip seed that preserves alignment-only style state after reopen |
| `xlsx-roundtrip` | `apply_style_formatting_depth_roundtrip_success.bin` | successful round-trip seed that preserves formatting-depth style state after reopen |

Matching promotion metadata lives under:

```text
jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata/
```

Those metadata files document:
- source artifact path
- promoted input path
- replay outcome kind
- replay text artifact path

---

## Present Strengths

The current Jazzer layer is strongest at:
- ordered operation and command interaction coverage
- validation-boundary discovery for sheet names, ranges, addresses, widths, heights, and freeze
  pane coordinates
- a committed custom seed floor that now includes readable public example requests and replay-
  verified binary workflow seeds
- style-aware `.xlsx` round-trip verification for formatting-depth features
- explicit source/persistence-mode telemetry for service-level workflow fuzzing
- deterministic tests that protect the Jazzer reporting and verifier infrastructure itself
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
- data-validation, conditional-formatting, table, chart, picture, and pivot coverage
- extremely large-workbook or streaming-strategy fuzzing
- cross-run corpus-health scoring beyond counts and newest-entry inspection
- automatic conversion of promoted inputs into deterministic main-suite tests
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
