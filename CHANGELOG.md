# Changelog

Notable changes to this project are documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.30.0] - 2026-04-07

### Added

- Style-system expansion: `APPLY_STYLE` and cell-style reads now expose a nested style contract
  with `numberFormat`, `alignment`, `font`, `fill`, `border`, and `protection` groups. The
  shipped style surface now includes text rotation, indentation, cell `locked` and
  `hiddenFormula` flags, per-side border colors, and patterned fills with foreground or
  background colors.
- Rich-text cell authoring: `SET_CELL`, `SET_RANGE`, and `APPEND_ROW` now accept typed
  `RICH_TEXT` values with ordered runs and optional per-run font overrides.
- String cell reads in `GET_CELLS`, `GET_WINDOW`, and `GET_SHEET_SCHEMA` now surface optional
  structured `richText` runs alongside `stringValue`, so authored rich text round-trips through
  the existing cell-introspection surface instead of a separate read family.

### Changed

- Public request examples, README snippets, and committed Jazzer protocol-request seeds now use
  the nested `APPLY_STYLE` JSON shape instead of the old flat style fields, so docs, examples,
  and deterministic replay all reflect the live public contract.
- Jazzer style generation, style-kind coverage telemetry, protocol response invariants, and
  `.xlsx` reopen verification now operate on the full nested style model rather than a flat
  subset.
- Jazzer typed-value generation, protocol invariants, and `.xlsx` reopen verification now assert
  rich-text persistence explicitly, including run ordering, non-empty run text, concatenation
  back to `stringValue`, and effective per-run font facts on read-back.
- Border color patches now require an effective visible border style on the same side, either set
  directly or inherited from `border.all`, instead of tolerating color-only states that Excel does
  not model as a visible border.

### Fixed

- `APPLY_STYLE` border patches that clear a side or `border.all` back to `NONE` no longer crash on
  workbooks whose underlying border XML already has no stored color entry. Border-color clearing
  now avoids POI's unsafe unset path when there is no color to remove, and the case is locked into
  deterministic `.xlsx` round-trip regression coverage.

## [0.29.0] - 2026-04-07

### Added

- Column structural editing:
  `INSERT_ROWS`, `DELETE_ROWS`, `SHIFT_ROWS`, `INSERT_COLUMNS`, `DELETE_COLUMNS`,
  `SHIFT_COLUMNS`, `SET_ROW_VISIBILITY`, `SET_COLUMN_VISIBILITY`, `GROUP_ROWS`,
  `UNGROUP_ROWS`, `GROUP_COLUMNS`, and `UNGROUP_COLUMNS`.
- `GET_SHEET_LAYOUT` now reports row and column `hidden`, `outlineLevel`, and `collapsed`
  facts alongside explicit size so grouped or hidden structure can be read back cleanly.
- New example request: `examples/row-column-structure-request.json`.
- Round-trip and Jazzer verification now assert persisted row or column layout state for column structural
  operations, including hidden-state shifts, grouping, and save-reopen fidelity.
- Jazzer promoted seed `row_column_structure_request` now covers the public structural-edit
  request surface, including row or column insertion, deletion, shifting, visibility, grouping,
  and `GET_SHEET_LAYOUT` reads.

### Changed

- `./check.sh` now emits `[CHECK-PULSE]` stage progress lines during long-running verification
  and captures `[CHECK-DIAG]` stall snapshots with process-tree, `lsof`, thread-dump, and log-tail
  artifacts when a stage stops making semantic progress. Diagnosed stalled stages are now
  terminated with a stable failure instead of waiting indefinitely.
- Nested Jazzer verification now emits stable `[JAZZER-PULSE]` progress lines for deterministic
  support tests, per-harness regression targets, per-input committed-seed replay, and standalone
  harness replay, so long-running Stage 2 verification exposes real progress instead of appearing
  silent.
- Promoted-input semantic replay now runs only in the dedicated Jazzer regression runner tasks
  instead of inside `PromotionMetadataTest`, keeping the support-test JVM structural-only while
  preserving committed replay-contract verification in the isolated tool runtime.
- Deterministic Jazzer replay now uses a project-owned pure-Java scalar fuzz-data cursor for the
  subset of `FuzzedDataProvider` behavior GridGrind actually consumes, so `jazzer/bin/replay`
  and committed-seed regression no longer depend on Jazzer's native replay provider loading in a
  fresh JVM.
- Column structural edits now re-normalize explicit column metadata after insert, delete, and
  shift operations so hidden state, outline state, and other explicit column-definition facts move
  with the authored columns instead of being left behind in stale XML.
- GridGrind now documents and surfaces two product-owned structural-edit limits:
  `LIM-016` rejects edits that would move or truncate tables, sheet-owned autofilters, or data
  validations; `LIM-017` rejects column structural edits when formulas or formula-defined names
  are present because Apache POI leaves some column references stale.

### Fixed

- Row and column destructive structural edits now reject range-backed named ranges before Apache
  POI can rewrite them into broken `#REF!` formulas. `DELETE_ROWS`, `SHIFT_ROWS`,
  `DELETE_COLUMNS`, and `SHIFT_COLUMNS` now preserve workbook integrity by allowing only safe
  named-range cases such as full-band moves and completely untouched ranges.
- Column structural edits now preserve Apache POI's required `<cols>` worksheet container even
  when a sheet has no explicit column metadata, so follow-up operations like `GROUP_COLUMNS`,
  `UNGROUP_COLUMNS`, and column visibility changes no longer crash after insert, delete, or shift
  edits rebuild column definitions.
- Collapsed row groups at the sparse tail of a sheet now persist their control-row marker across
  save and reopen, so `GROUP_ROWS(..., collapsed=true)` round-trips the exposed `collapsed` fact
  in `GET_SHEET_LAYOUT` instead of losing it on reopen.
- Sparse rows touched by `UNGROUP_ROWS` now normalize Apache POI's internal outline-level
  sentinel `-1` back to public outline level `0`, so row layout reads and round-trip assertions
  stay stable after save and reopen.
- Promoted Jazzer metadata and replay text now store project-relative paths instead of
  hard-coded workspace paths. `jazzer/bin/promote`, promoted-metadata refresh, and orphaned-seed
  detection now resolve those paths against the project directory, so replay validation works
  correctly from alternate worktrees.

## [0.28.0] - 2026-04-06

### Changed

- `GridGrindJson.cleanJacksonMessage` now strips Jackson configuration-advice suffixes of the
  form `(set X.Y to 'Z' to allow)` in addition to the existing source-location, subtype, and
  POJO-property noise patterns. This closes a class of raw-message leaks: any null-into-primitive
  coercion on any field type now has its configuration advice stripped before the message reaches
  `productOwnedJacksonMessage`, regardless of whether a dedicated dispatch arm exists for that
  exact exception shape.

### Fixed

- `GridGrindJson` now surfaces a clean `Missing required field '<name>'` error when a required
  primitive boolean field (e.g. `stopIfTrue`) is supplied as JSON `null`. Previously the raw
  Jackson internal configuration-advice message
  (`"set DeserializationConfig.DeserializationFeature.FAIL_ON_NULL_FOR_PRIMITIVES to 'false'
  to allow"`) leaked directly into the `InvalidRequestShapeException` message seen by agents.
- Protocol catalog `DifferentialStyleInput` descriptor now lists all nine optional fields
  (`numberFormat`, `bold`, `italic`, `fontHeight`, `fontColor`, `underline`, `strikeout`,
  `fillColor`, `border`). Previously the field list was empty, so `--describe-type` showed no
  fields at all for this type.
- Protocol catalog descriptions for `FORMULA` cell input, `FORMULA_LIST` and `CUSTOM_FORMULA`
  data-validation rule types, and `FORMULA_RULE` conditional-formatting rule type no longer
  instruct agents to omit the leading `=` sign. GridGrind accepts and strips it automatically;
  the prior wording contradicted the actual behavior observed when submitting `=SUM(...)`.

### Added

- `PromotionMetadataTest.everyInputFileHasPromotionMetadata` asserts that every file committed
  to any harness input directory has a corresponding promoted-metadata entry. Previously only
  the inverse direction was enforced: metadata entries were replayed but no test checked whether
  every input *file* had metadata. A seed hand-dropped into the input directory without running
  `jazzer/bin/promote` would compile, replay in regression mode, but leave no stable contract
  that `PromotionMetadataTest` could assert against. The new test closes that gap for all four
  harnesses simultaneously.
- `JazzerReportSupport.orphanedInputs` and `promotedInputPaths` — new public methods that
  identify input files with no promoted-metadata entry by building a set of all
  `promotedInputPath` values from the metadata tree and diffing against the input directory.
  Used by the new test and by `list-corpus`.
- `list-corpus` (`jazzer/bin/list-corpus`) now surfaces a `WARNING: Seeds Without Promotion
  Metadata` section for any harness that has orphaned input files, listing each file and
  reminding the operator to run `jazzer/bin/promote`. Previously the gap was invisible from
  the operator tooling.
- `GridGrindJsonTest` integration test `wrapsNullPrimitiveBooleanFieldAsInvalidRequestShapeWithFieldName`
  verifies that a full `readRequest` round-trip with `"stopIfTrue": null` produces
  `InvalidRequestShapeException` with message `"Missing required field 'stopIfTrue'"` and no
  Jackson internals in the message.
- Four `GridGrindJsonTest` unit tests for the `mismatchedInputMessage` null-into-primitive paths
  (named property, no path, array-index path) and `cleanJacksonMessage` configuration-advice
  stripping.
- `GridGrindJsonTest` unit test `mismatchedInputMessageWithNullOriginalMessageReturnsFallback`
  covering the `original == null` branch of `mismatchedInputMessage`.
- Jazzer promoted seed `invalid_request_shape_null_primitive_boolean` covering the
  `SET_CONDITIONAL_FORMATTING` request with `"stopIfTrue": null`, captured as
  `EXPECTED_INVALID` / `INVALID_REQUEST_SHAPE`.
- Jazzer promoted seed `invalid_request_shape_null_primitive_int` covering an `APPLY_STYLE`
  request with `"fontHeight": {"type": "TWIPS", "twips": null}`, captured as `EXPECTED_INVALID`
  / `INVALID_REQUEST_SHAPE` with message `"Missing required field 'twips'"`. Proves that the
  `cleanJacksonMessage` configuration-advice stripping generalizes to non-boolean primitives.
- Fourteen previously unpromoted protocol-request seeds retroactively promoted with full
  expectation metadata and replay text: `clear_on_empty_cells`, `delete_last_sheet`,
  `duplicate_request_id`, `formula_equals_prefix`, `get_cells_invalid_address`,
  `get_cells_out_of_bounds_address`, `get_window_overflow`, `introspection_analysis_request`,
  `invalid_email_no_at_sign`, `schema_empty_sheet`, `schema_formula_cells`,
  `sheet_name_too_long`, `unknown_field_rejection`, `window_size_limit_exceeded`. All were
  curated seeds in the corpus without regression contracts; `PromotionMetadataTest` now covers
  them.

## [0.27.0] - 2026-04-05

### Changed

- Eleven shadow enums deleted from `protocol.dto` (`HorizontalAlignment`, `VerticalAlignment`,
  `BorderStyle`, `SheetVisibility`, `PrintOrientation`, `PaneRegion`, `ComparisonOperator`,
  `DataValidationErrorStyle`, `ConditionalFormattingIconSet`, `ConditionalFormattingThresholdType`,
  `ConditionalFormattingUnsupportedFeature`). All protocol DTOs, reports, and operation types now
  reference the canonical `Excel*` engine enums directly via the existing module dependency.
- `DefaultGridGrindRequestExecutor` (2 571 lines) decomposed into three package-private converter
  classes: `WorkbookCommandConverter` (protocol operations to engine commands),
  `WorkbookReadCommandConverter` (protocol read operations to engine read commands), and
  `WorkbookReadResultConverter` (engine read results to protocol reports). The executor now
  delegates to these converters and contains no conversion logic. Each converter exposes
  package-private static methods for direct unit testing without reflection.

### Added

- New Jazzer protocol-request regression seed `conditional_formatting_request` covering
  `SET_CONDITIONAL_FORMATTING` (formula rule, cell-value rule with `LESS_THAN`, cell-value rule
  with `BETWEEN` including a differential border), `CLEAR_CONDITIONAL_FORMATTING` with a selected
  range, `GET_CONDITIONAL_FORMATTING`, and `ANALYZE_CONDITIONAL_FORMATTING_HEALTH`. No seed
  exercising conditional-formatting operations existed previously.

### Fixed

- Promoted protocol-request seed `conditional_formatting_request` uses `conditionalFormatting`
  for `SET_CONDITIONAL_FORMATTING` payloads, matching `WorkbookOperation.SetConditionalFormatting`
  JSON and Jazzer promotion-metadata replay expectations.

## [0.26.0] - 2026-04-04

### Added

- New Jazzer protocol-request regression seed `pane_and_print_reset_request` covering the
  previously untested pane and print-layout variants: `PaneInput.NONE` (reset), `PaneInput.SPLIT`
  (x/y offsets with active-pane region), `PrintAreaInput.NONE`, `PrintScalingInput.AUTOMATIC`,
  and `CLEAR_PRINT_LAYOUT`. The existing `structural_layout_request` seed only covered `FROZEN`
  panes with `FIT` scaling.
- Protocol catalog now publishes `paneTypes`, `printAreaTypes`, `printScalingTypes`,
  `printTitleRowsTypes`, and `printTitleColumnsTypes` as nested type groups, and
  `headerFooterTextInputType` and `printLayoutInputType` as plain type groups. Agents using
  `--print-protocol-catalog` or `--describe-operation SET_SHEET_PANE` / `SET_PRINT_LAYOUT`
  previously received dangling `NESTED_TYPE_GROUP` or `PLAIN_TYPE_GROUP` references with no
  matching catalog entry; those entries are now present and complete.
- Bidirectional validation between the field-shape group maps in `CatalogFieldMetadataSupport`
  and the descriptor lists in `GridGrindProtocolCatalog`: a type registered in the field-shape
  map with no corresponding catalog descriptor now raises `IllegalStateException` at startup
  rather than silently producing an incomplete catalog.

### Changed

- Eight public nested types (`Catalog`, `TypeEntry`, `FieldEntry`, `NestedTypeGroup`,
  `PlainTypeGroup`, `FieldShape`, `FieldRequirement`, `ScalarType`) extracted from
  `GridGrindProtocolCatalog` to individual top-level files in the `catalog` package.
  Wire format, catalog content, and public API are unchanged.
- Protocol-catalog construction now uses a small internal gatherer seam for ordered uniqueness
  and reflected field-metadata expansion instead of hand-rolled duplicate and ordering logic
  inside `GridGrindProtocolCatalog`.
- Built-in discovery output remains deterministic and contract-identical, while request-template
  generation intentionally stays a plain constant because it does not warrant gatherer-based
  abstraction.
- Release and container workflows now support tag-targeted `workflow_dispatch` reruns, and the
  release procedure now codifies protected-`main` CI requirements plus automatic release-branch
  cleanup.
- GHCR publication verification now runs both `docker pull` and `docker run` through the same
  anonymous Docker config, matching the operator release protocol instead of silently falling
  back to ambient credentials.

## [0.25.0] - 2026-04-03

### Changed

- GridGrind now compiles and verifies on the JPMS module path across `engine`, `protocol`, and
  `cli`, enforcing the intended `cli -> protocol -> engine` dependency graph in normal builds.
- The protocol implementation is now split into protocol-owned packages by responsibility:
  `dto`, `operation`, `read`, `catalog`, `exec`, and `json`, instead of the older flat
  package layout.
- `DefaultGridGrindRequestExecutor` is now the sole engine-aware class in protocol main source.
  Generic problem construction, JSON handling, and catalog generation remain protocol-owned.

## [0.24.0] - 2026-04-03

### Added

- Structural-layout public surface:
  `SET_SHEET_PANE`, `SET_SHEET_ZOOM`, `SET_PRINT_LAYOUT`, `CLEAR_PRINT_LAYOUT`, and
  `GET_PRINT_LAYOUT`.
- New engine seams for sheet-view and print-layout control, including explicit pane-state,
  zoom, print-area, orientation, fit scaling, repeating rows or columns, and plain
  header/footer text support.

### Changed

- `GET_SHEET_LAYOUT` now reports a generalized `pane` family plus effective `zoomPercent`
  instead of the older freeze-only layout contract.
- The machine-readable protocol catalog, public docs, and Jazzer readable seeds now describe the
  generalized pane and print-layout contract rather than the superseded `FREEZE_PANES` shape.

### Fixed

- Jazzer read introspection and `.xlsx` round-trip structural invariants now treat pane state as
  a generalized workbook-view concern instead of a freeze-only special case.

## [0.23.0] - 2026-04-02

### Added

- Sheet-management public surface:
  `COPY_SHEET`, `SET_ACTIVE_SHEET`, `SET_SELECTED_SHEETS`, `SET_SHEET_VISIBILITY`,
  `SET_SHEET_PROTECTION`, and `CLEAR_SHEET_PROTECTION`.
- `GET_WORKBOOK_SUMMARY` now reports the typed `EMPTY` versus `WITH_SHEETS` workbook-summary
  shape, and non-empty workbooks expose `activeSheetName` plus `selectedSheetNames`.
- `GET_SHEET_SUMMARY` now reports `visibility` and typed sheet-protection state alongside the
  existing structural row and column facts.
- New public example `examples/sheet-management-request.json` covering sheet copy, active and
  selected sheet state, visibility, protection, workbook summary, and sheet summary reads.

### Changed

- Sheet-copy execution is now GridGrind-owned instead of delegating to Apache POI's raw
  `cloneSheet()` behavior. Copying preserves supported sheet-local content while rejecting
  unsupported copy cases such as tables and sheet-scoped formula-defined named ranges.
- Jazzer sequence generation now uses a stable byte-selector grammar for workflow, command, and
  read-family dispatch, so expanding the authored surface no longer requires mutating bounded
  selector ranges in place.
- The committed `sheet_management_request` Jazzer protocol-request seed now exercises the shipped
  sheet-state contract instead of the older rename and move only slice.

### Fixed

- `DELETE_SHEET` now shares the same visible-sheet invariant as `SET_SHEET_VISIBILITY`, so
  deleting the last visible sheet returns `INVALID_REQUEST` instead of crashing during workbook
  view-state normalization.
- Workbook and sheet summaries now expose active-sheet, selected-sheet, visibility, and
  protection state through the public response model instead of truncating the new facts at
  the engine boundary.
- Sheet-management copy and workbook-view normalization now have direct round-trip and fuzz-backed
  verification, including empty-sheet copies, protected-sheet copies, and invalid active-tab
  repair paths.
- `CLEAR_SHEET_PROTECTION` is now idempotent on already unprotected sheets instead of delegating
  into an Apache POI removal path that could throw during `.xlsx` round-trip fuzzing.
- Jazzer promotion metadata is now refreshed against the stable selector grammar, so replay
  expectations remain truthful after the sheet-management expansion instead of silently
  describing a previous generator contract.

## [0.22.0] - 2026-04-02

### Added

- Conditional formatting public surface:
  `SET_CONDITIONAL_FORMATTING`, `CLEAR_CONDITIONAL_FORMATTING`,
  `GET_CONDITIONAL_FORMATTING`, and `ANALYZE_CONDITIONAL_FORMATTING_HEALTH`.
- New public example `examples/conditional-formatting-request.json` covering block authoring,
  factual reads, conditional-formatting health, and aggregate workbook findings.

### Changed

- `ANALYZE_WORKBOOK_FINDINGS` now aggregates conditional-formatting findings alongside formula,
  data-validation, autofilter, table, hyperlink, and named-range findings.

### Fixed

- Protocol coverage, executor mapping coverage, and `.xlsx` round-trip Jazzer invariants now
  assert conditional-formatting authoring and persistence instead of leaving the new family
  under-verified.
- Table header mutations made after `SET_TABLE`, including header-range style patches that change
  typed header display text, now synchronize the persisted table-column metadata immediately.
  `ExcelWorkbook.save()` also normalizes every table header before persistence as a backstop, so
  `GET_TABLES`, table health analysis, save/reopen behavior, and `.xlsx` round-trip fuzzing all
  observe one converged table-header state instead of drifting between sheet cells and table XML.
- Jazzer now has promoted `.xlsx` round-trip success seeds covering both the former table-header
  rewrite crash and the later typed-header style-display crash, and the deterministic round-trip
  verifier suite asserts that header rewrites, header clears, and header-range style changes
  survive save and reopen without metadata drift.

## [0.21.0] - 2026-04-01

### Added

- Table and autofilter public surface:
  `SET_AUTOFILTER`, `CLEAR_AUTOFILTER`, `SET_TABLE`, `DELETE_TABLE`, `GET_AUTOFILTERS`,
  `GET_TABLES`, `ANALYZE_AUTOFILTER_HEALTH`, and `ANALYZE_TABLE_HEALTH`.
- New public example `examples/table-autofilter-request.json` covering sheet-level autofilters,
  table authoring, factual reads, and both health-analysis families.
- Jazzer now has promoted seeds for a readable table-plus-autofilter request, a protocol
  workflow dominated by autofilter behavior, and an `.xlsx` round-trip invalid-table case.
- Jazzer promotion metadata now carries a stable replay expectation contract, and the new
  `jazzer/bin/refresh-promoted-metadata` command refreshes committed replay metadata after
  intentional generator or replay-shape changes.

### Changed

- `ANALYZE_WORKBOOK_FINDINGS` now aggregates autofilter and table findings alongside formula,
  data-validation, hyperlink, and named-range findings.

### Fixed

- Table and autofilter logic now share a dedicated sheet-structure support seam instead of
  coupling controllers through package-private helper methods.
- Jazzer promoted-seed verification no longer drifts silently when sequence-generation behavior
  changes; deterministic support tests now replay every promoted metadata entry and assert that
  its stored replay expectation still matches reality.

## [0.20.0] - 2026-04-01

### Added

- Data validation public surface:
  `SET_DATA_VALIDATION`, `CLEAR_DATA_VALIDATIONS`, `GET_DATA_VALIDATIONS`, and
  `ANALYZE_DATA_VALIDATION_HEALTH`.
- New public example `examples/data-validation-request.json` covering validation authoring,
  partial clearing, factual reads, and health analysis.
- Jazzer now has a promoted valid validation workflow seed and a promoted expected-invalid
  validation seed, so the committed regression floor covers both supported and rejected
  validation shapes.

### Changed

- `ANALYZE_WORKBOOK_FINDINGS` now aggregates data-validation findings alongside formula,
  hyperlink, and named-range findings.
- Protocol discovery now publishes `rangeSelectionTypes`, `dataValidationRuleTypes`,
  `dataValidationInputType`, `dataValidationPromptInputType`, and
  `dataValidationErrorAlertInputType`, so black-box consumers can author the full validation
  workflow without shape inference.
- Developer docs, the Jazzer coverage inventory, and the Apache POI parity inventory now treat
  data validation as shipped behavior instead of a planned gap.

### Fixed

- `.xlsx` round-trip verification, executor integration coverage, and Jazzer invariants now assert
  that normalized data-validation state survives save and reopen instead of silently ignoring the
  new command family.
- `.xlsx` round-trip verification no longer duplicates stale command-semantics assumptions for
  style and metadata persistence. The verifier now snapshots expected pre-save workbook state from
  the actual in-memory workbook, so `APPEND_ROW` date-time writes onto styled blank rows replay
  cleanly and stay covered by the committed Jazzer seed floor.
- `GET_DATA_VALIDATIONS` now exposes only observable public states: `SUPPORTED` and
  `UNSUPPORTED`. Invalid workbook validation structures that Apache POI refuses to materialize are
  no longer advertised as a separate `MALFORMED` entry family, so the read contract matches the
  real workbook-loading seam.

## [0.19.0] - 2026-04-01

### Changed

- `--print-protocol-catalog` now publishes field descriptors instead of loose field-name lists.
  Every catalog entry states whether a field is required or optional and exposes the exact scalar,
  list, nested-group, or plain-group shape accepted by that field.
- The `Release` and `Container` workflows now serialize publication per workflow and tag ref,
  verify the external GitHub release and GHCR handoff after publishing, and treat publication as
  a converged public state instead of a one-shot side effect.
- `./check.sh` now syntax-checks the release-surface shell scripts before the Docker smoke stage,
  so publication helpers fail fast locally and in CI when a shell edit breaks the release path.

### Fixed

- CLI and docs discovery guidance now explain that the machine-readable protocol catalog is the
  authoritative black-box contract for field requirements and polymorphic field shapes, so
  operations such as `SET_HYPERLINK`, `SET_RANGE`, and the selection-based reads can be authored
  without inference.
- Duplicate tag-triggered release runs no longer fail spuriously with `Release.tag_name already
  exists`; GitHub Release publication is now idempotent and asset-safe under duplicate dispatch.
- Duplicate or delayed tag-triggered container publication now has a built-in pull-and-run
  verification step, so the workflow confirms the exact version tag and `latest` are both
  publicly runnable before cleanup proceeds.

## [0.18.0] - 2026-03-31

### Added

- `./check.sh` now runs a fourth stage, `scripts/docker-smoke.sh`, which builds the local Docker
  image and verifies `--help`, `--version`, request-file loading, response-file writing, and
  `.xlsx` persistence from a non-default working directory using weird path names.
- CI now includes a separate `Docker smoke` job, and the `Container` workflow runs the same smoke
  script before publishing multi-arch images to GHCR.

### Changed

- `FILE` hyperlinks continue the hard-break path contract: requests use `FILE.path`, plain paths
  and `file:` URIs normalize to plain paths, read surfaces return plain paths, and hyperlink
  health resolves relative file targets against the workbook location.
- `APPEND_ROW` continues to use value-bearing row semantics, so metadata-only rows do not shift
  the append cursor.
- `AUTO_SIZE_COLUMNS` continues to use deterministic content-based sizing so container and local
  runs agree on column widths.
- CLI `--help`, protocol catalog discovery, `README.md`, and the public reference docs now
  explain the file workflow explicitly: stdin vs `--request`, stdout vs `--response`, `SAVE_AS`
  vs `OVERWRITE`, current-working-directory path resolution, and Docker `-w /workdir` usage.

### Fixed

- The Docker image remains workdir-independent: `docker run -w /workdir ... --help` and
  file-backed request/response flows now have automated regression coverage.
- Hyperlink health continues to report missing local file targets instead of silently treating
  them as healthy.
- Invalid request-shape and invalid cell-address failures continue to expose product-owned
  diagnostics instead of parser, POI, or Java implementation detail.
- Typed value writes continue to preserve existing style, hyperlink, and comment state, and
  `DATE` / `DATE_TIME` writes continue to merge their formats onto the existing style instead of
  replacing it.

## [0.17.0] - 2026-03-31

### Fixed

- Response-side hyperlink payloads now reuse the canonical discriminated hyperlink shape used by
  `SET_HYPERLINK`. `GET_HYPERLINKS`, `GET_CELLS`, and `GET_WINDOW` now return `FILE` targets in
  the `path` field instead of leaking the legacy `target` field on read.
- Failure-response `causes` entries now expose stable GridGrind problem codes and product-owned
  messages instead of raw Java exception class names or parser-library internals. This is a
  **wire format breaking change** for any client that was matching on the old `type` or
  `className` fields.
- `INVALID_REQUEST_SHAPE` now returns product-owned messages for missing required fields as well,
  so shape failures no longer leak parser or Java type metadata through that branch either.
- Typed value writes now preserve existing cell style, hyperlink, and comment state instead of
  resetting presentation when `SET_CELL`, `SET_RANGE`, or `APPEND_ROW` overwrites a styled blank
  cell.
- `DATE` and `DATE_TIME` writes now merge their required number formats onto the existing cell
  style instead of replacing fill, border, font, alignment, or wrap state.
- Jazzer and deterministic round-trip coverage now assert style preservation when `APPEND_ROW`
  reuses styled blank rows under value-bearing append semantics.

## [0.16.0] - 2026-03-31

### Fixed

- The Docker image entrypoint now uses an absolute JAR path, so `docker run ... -w /any/path`
  works reliably instead of failing before GridGrind starts with `Unable to access jarfile`.
- `ANALYZE_HYPERLINK_HEALTH` now reports missing local file targets and unresolved relative file
  targets instead of silently treating those cases as healthy.
- `INVALID_REQUEST_SHAPE` messages are now product-owned and concise. Unknown fields, unknown
  type values, and wrong token shapes no longer leak Jackson or Java class names into the public
  response.

### Changed

- `FILE` hyperlink targets are now written with the field name `path` instead of `target`. The
  write contract accepts either plain file paths or `file:` URIs, and all read surfaces return
  normalized plain path strings. This is a **wire format breaking change** for clients that still
  send or expect `target` on `FILE` hyperlinks.
- `APPEND_ROW` now appends after the last value-bearing row. Rows that contain only style,
  comment, or hyperlink metadata no longer shift the append cursor.
- `AUTO_SIZE_COLUMNS` now uses deterministic content-based sizing instead of host font metrics,
  so headless, Docker, and local runs produce the same column widths.
- `./check.sh` now runs nested Jazzer `check` after the root quality gates and before CLI fat-JAR
  packaging, giving the one-command local gate deterministic support-test and committed-seed
  regression coverage as well.

### Added

- New public example `examples/file-hyperlink-health-request.json` showcasing `FILE.path`,
  `file:` URI normalization, hyperlink metadata reads, and hyperlink-health analysis.

## [0.15.0] - 2026-03-31

### Fixed

- Number cells in `GET_CELLS` and `GET_WINDOW` responses now return `declaredType` and
  `effectiveType` as `NUMBER` instead of `NUMERIC`. This is a **wire format breaking change**:
  any client that matched on `"NUMERIC"` must be updated to `"NUMBER"`. The value `NUMERIC` was
  an Apache POI internal enum name leaking into the wire vocabulary; `NUMBER` matches the input
  side and is self-explanatory.
- `SET_CELL` and `SET_RANGE` now accept a leading `=` in `FORMULA` cell input. Previously, a
  formula like `"=SUM(A1:A3)"` was sent to Apache POI with the `=` retained, causing an
  `InvalidFormulaException`. The leading `=` is now stripped automatically before handing the
  expression to the engine.
- `DELETE_SHEET` now returns `INVALID_REQUEST` when the operation would delete the last remaining
  sheet in the workbook. Previously the request was forwarded to Apache POI, which threw an
  unclassified `IllegalStateException`; the error is now proactively detected and surfaced with a
  clear message before any workbook state is modified.
- Validation error messages for `SET_ROW_HEIGHT` and `SET_COLUMN_WIDTH` now include both the
  enforced limit and the supplied value so the caller can identify the violation without
  re-reading the protocol catalog. For example: `"heightPoints must not exceed 1638.35 (Excel
  storage limit: 32767 twips): got 2000.0"` instead of `"must be less than or equal to 1638.35"`.

### Changed

- `--print-protocol-catalog` now accepts an optional `--operation <id>` flag. When supplied, the
  output contains only the single catalog entry matching that operation ID. Unknown IDs produce
  an `INVALID_ARGUMENTS` error. Without `--operation`, the full catalog is returned unchanged.
- `APPLY_STYLE` catalog summary now documents the write vs. read shape asymmetry for borders:
  the `border` write object (with `all`, `top`, `right`, `bottom`, `left` sub-objects) is not
  mirrored in the cell snapshot; read responses use flat top-level fields `topBorderStyle`,
  `rightBorderStyle`, `bottomBorderStyle`, and `leftBorderStyle`.
- `DELETE_SHEET` catalog summary now states that deleting the last sheet in a workbook returns
  `INVALID_REQUEST`.
- `DATE` and `DATE_TIME` cell input summaries in `--print-protocol-catalog` and `--help` now
  correctly state that `GET_CELLS` returns `declaredType=NUMBER` (not `declaredType=NUMERIC`) for
  these cells. The help text `Limits:` section is updated accordingly.

## [0.14.0] - 2026-03-30

### Fixed

- `GET_CELLS` and `GET_WINDOW` now reject cell addresses that exceed the Excel 2007 sheet
  boundary (row > 1,048,575 or column > 16,383). Previously, addresses like `XFE1` (column
  16,384) were accepted by `CellReference` with non-negative indices and returned a blank
  snapshot instead of failing with `INVALID_CELL_ADDRESS`.
- `GET_WINDOW` now rejects window dimensions that would extend the window beyond the Excel
  2007 sheet boundary. Previously, a window starting at a valid address could silently overflow
  if `topLeft.row + rowCount - 1 > 1,048,575` or `topLeft.col + columnCount - 1 > 16,383`.
- Row height read-back from `GET_SHEET_LAYOUT` now returns an exact value for heights stored as
  twips. Previously, `row.getHeightInPoints()` returned a `float` which introduced floating-point
  imprecision (e.g., 1,638.35 points stored as 32,767 twips read back as 1,638.3499755859375).
  Heights are now read as integer twips divided by 20.0, eliminating the imprecision.
- Error messages from Jackson for unknown `type` discriminators no longer include internal
  fully-qualified class names (e.g., `dev.erst.gridgrind.protocol.WorkbookReadOperation`) or
  Jackson-internal POJO property annotations. The message now contains only the unknown
  discriminator value and the list of known type IDs.
- `MOVE_SHEET` error message for an out-of-range `targetIndex` now clearly states the workbook's
  sheet count and the valid 0-based index range. Previously the message said
  "between 0 and N (inclusive)" without clarifying what N represented.

### Changed

- `GET_SHEET_SCHEMA` now counts formula cells by their evaluated result type (NUMERIC, STRING,
  BOOLEAN, ERROR) in `observedTypes` and `dominantType`, rather than as FORMULA. This makes
  the schema reflect the data a consumer actually reads, not the cell's internal storage type.

## [0.13.0] - 2026-03-30

### Fixed

- `GET_CELLS` no longer silently returns a blank cell snapshot for a malformed address (e.g.
  `BADADDR`, `A0`). Requests containing any address that Apache POI cannot resolve to valid row
  and column indices now fail immediately with `INVALID_CELL_ADDRESS`. Previously, `CellReference`
  returned row index `-1` for unparseable addresses, and the engine treated that as an absent cell
  and returned a `BLANK` snapshot with no error.
- Row height validation now enforces the exact documented boundary of 1,638.35 points
  (32,767 twips). Previously, `Math.round`-based twips conversion accepted values up to
  approximately 1,638.37 because `Math.round(1638.37 × 20) = 32767`. The validation now uses a
  direct floating-point comparison `heightPoints > Short.MAX_VALUE / 20.0` so that any value above
  1,638.35 is rejected regardless of rounding.

### Changed

- All response `type` discriminator values now echo the corresponding request `type` exactly. This
  is a **wire format breaking change** for any client that inspects the `type` field in read results
  or persistence outcomes. The full mapping from old to new discriminator values:

  Read result discriminators:
  - `WORKBOOK_SUMMARY` → `GET_WORKBOOK_SUMMARY`
  - `NAMED_RANGES` → `GET_NAMED_RANGES`
  - `SHEET_SUMMARY` → `GET_SHEET_SUMMARY`
  - `CELLS` → `GET_CELLS`
  - `WINDOW` → `GET_WINDOW`
  - `MERGED_REGIONS` → `GET_MERGED_REGIONS`
  - `HYPERLINKS` → `GET_HYPERLINKS`
  - `COMMENTS` → `GET_COMMENTS`
  - `SHEET_LAYOUT` → `GET_SHEET_LAYOUT`
  - `FORMULA_SURFACE` → `GET_FORMULA_SURFACE`
  - `SHEET_SCHEMA` → `GET_SHEET_SCHEMA`
  - `NAMED_RANGE_SURFACE` → `GET_NAMED_RANGE_SURFACE`
  - `FORMULA_HEALTH` → `ANALYZE_FORMULA_HEALTH`
  - `HYPERLINK_HEALTH` → `ANALYZE_HYPERLINK_HEALTH`
  - `NAMED_RANGE_HEALTH` → `ANALYZE_NAMED_RANGE_HEALTH`
  - `WORKBOOK_FINDINGS` → `ANALYZE_WORKBOOK_FINDINGS`

  Persistence outcome discriminators:
  - `NOT_SAVED` → `NONE`
  - `SAVED_AS` → `SAVE_AS`
  - `OVERWRITTEN` → `OVERWRITE`

  With symmetric naming, the response `type` field directly identifies which read or persistence
  operation produced it, eliminating any need for a client-side translation table.
- `GET_WORKBOOK_SUMMARY` catalog summary now states that the response includes the list of sheet
  names in the workbook.
- `GET_CELLS` catalog summary now documents that invalid cell addresses (malformed or out-of-range)
  produce `INVALID_CELL_ADDRESS`, and that `effectiveType` is always `FORMULA` for formula cells
  regardless of the evaluated result type.
- `GET_WINDOW` catalog summary now notes that `effectiveType` is always `FORMULA` for formula
  cells.
- Persistence outcome summaries in `--print-protocol-catalog` now state that the response `type`
  echoes the request persistence `type` exactly (e.g. `NONE` persistence returns `type: "NONE"`).
- The container cleanup workflow now prunes GHCR package versions through GitHub's Packages API
  via `gh api`, anchored to the five newest tagged releases, instead of using the stale
  `actions/delete-package-versions` wrapper. This removes the Node20 deprecation warning and
  keeps complete multi-arch release groups together even when GitHub emits multiple untagged
  platform and attestation manifests per release.

## [0.12.0] - 2026-03-29

### Fixed

- `SAVE_AS` and `OVERWRITE` persistence now normalize the save path to its absolute canonical
  form before writing. Paths containing `..` segments (e.g. `/workdir/../out.xlsx`) are resolved
  to their canonical equivalents (`/out.xlsx`) before the file is written. `executionPath` in
  the response now always reflects the true path on disk. Previously, `..` segments were
  preserved in `executionPath` and the file was written to the un-normalized location.
- `GET_WINDOW` and `GET_SHEET_SCHEMA` now reject requests where `rowCount * columnCount` exceeds
  250,000 cells with `INVALID_REQUEST` before any workbook work occurs. Previously, large windows
  (e.g., 1000x1000) could crash the process with `OutOfMemoryError` and produce an empty response
  file. The 250,000-cell limit is a GridGrind operational constraint (not an Excel or Apache POI
  limit; Excel supports up to 1,048,576 rows x 16,384 columns) calibrated to prevent heap
  exhaustion during JSON response serialization in bounded-heap container environments.
- `CLEAR_RANGE` is now a no-op on rows and cells that do not physically exist. Previously it
  materialized phantom rows and cells into the sheet, inflating `physicalRowCount` and distorting
  `GET_SHEET_SUMMARY` results.
- `CLEAR_HYPERLINK` and `CLEAR_COMMENT` are now no-ops when the target cell does not physically
  exist, matching the idempotent behavior of `CLEAR_RANGE`. Previously they returned
  `CELL_NOT_FOUND`.
- `GET_SHEET_SCHEMA` now returns `dataRowCount = 0` when every cell in the inferred header row is
  blank. Previously it returned `rowCount - 1` even for empty sheets with no header data.
- `--request` and `--response` that resolve to the same path are now rejected at argument parse
  time with `INVALID_ARGUMENTS`. Previously the response write silently overwrote the request.

### Changed

- `Execution:` section of `--help` now reads "saves the workbook (unless persistence is NONE)"
  instead of the ambiguous word "persistence", which could be confused with the `persistence`
  JSON field.
- `--help` now includes a `Limits:` section listing all hard constraints upfront: `.xlsx`-only
  format, 31-character sheet names, 250,000-cell window cap, 255-unit column width ceiling,
  1,638-point row height ceiling, and the `DATE`/`DATE_TIME` write-only note. Agents and users
  can now read every hard constraint before constructing any request.
- `--help` now explicitly states that a NEW workbook starts with zero sheets and that
  `ENSURE_SHEET` must be used to create the first sheet.
- `ENSURE_SHEET` and `RENAME_SHEET` summaries in `--print-protocol-catalog` now state the
  31-character sheet name limit.
- `SET_COLUMN_WIDTH` summary in `--print-protocol-catalog` now states `widthCharacters` must
  be > 0 and ≤ 255.
- `SET_ROW_HEIGHT` summary in `--print-protocol-catalog` now states `heightPoints` must be > 0
  and ≤ 1,638.35 (32,767 twips).
- `DATE` and `DATE_TIME` cell input summaries in `--print-protocol-catalog` now note that these
  are write-only type hints stored as Excel serial numbers; `GET_CELLS` returns
  `declaredType=NUMERIC` with a formatted `displayValue`.
- `CLEAR_HYPERLINK` and `CLEAR_COMMENT` summaries in `--print-protocol-catalog` updated to
  reflect the no-op behavior on non-existent cells.
- `GET_WINDOW` and `GET_SHEET_SCHEMA` summaries in `--print-protocol-catalog` now state the
  250,000-cell limit.
- `GET_SHEET_SCHEMA` summary notes that `dataRowCount` is 0 when the header row is entirely blank.
- `NEW` source type summary now notes that a new workbook starts with zero sheets.
- `SAVE_AS` persistence type summary in `--print-protocol-catalog` now documents `requestedPath`
  (the literal path from the request) vs `executionPath` (the absolute normalized path where the
  file was written), and states that missing parent directories are created automatically.
- `GET_SHEET_SUMMARY` summary in `--print-protocol-catalog` now states the semantics of
  `physicalRowCount` (sparse materialized row count), `lastRowIndex` (0-based, -1 when empty),
  and `lastColumnIndex` (0-based, -1 when empty).
- `GET_CELLS` summary now documents the cell snapshot response shape: `address`, `declaredType`,
  `effectiveType`, `displayValue`, `style`, `metadata`, and type-specific value fields
  (`stringValue`, `numberValue`, `booleanValue`, `errorValue`, `formula`, `evaluation`).
- `GET_CELLS` summary now explicitly states that `style.fontHeight` in read responses is a plain
  object with both `twips` and `points` fields, not the discriminated `FontHeightInput` write
  format. `GET_WINDOW` summary cross-references the `GET_CELLS` cell snapshot shape.
- `fontHeightTypes` entries in `--print-protocol-catalog` now document the write format and the
  read-back shape asymmetry so agents can round-trip font height without format confusion.

### Added

- Protocol catalog `TypeEntry` now includes a `fieldEnumValues` map enumerating valid string
  values for fields that accept a finite enumerated set. `CellStyleInput.horizontalAlignment`
  lists all `ExcelHorizontalAlignment` values, `CellStyleInput.verticalAlignment` lists all
  `ExcelVerticalAlignment` values, and `CellBorderSideInput.style` lists all `ExcelBorderStyle`
  values. Agents can now discover valid alignment and border-style values from the catalog
  without trial-and-error.
- `docs/OPERATIONS.md` now includes a "Cell snapshot shape" subsection under `GET_CELLS`
  documenting all common and type-specific fields, the `fontHeight` read-vs-write asymmetry,
  and a field-level table for `GET_SHEET_SUMMARY` response semantics (`physicalRowCount`,
  `lastRowIndex`, `lastColumnIndex`). The `SAVE_AS` section now documents `requestedPath` vs
  `executionPath`.
- `docs/LIMITATIONS.md`: new reference document structured as a numbered registry (`LIM-001`
  through `LIM-015`). Each entry carries a stable ID, the enforced limit value, the error code
  and message raised on violation, the applicable operations, a code reference, and a UX
  reference. Covers the GridGrind operational window limit (250,000 cells), all Excel/Apache POI
  structural limits (rows, columns, text length, cell styles, hyperlinks, formula length,
  nested functions, function arguments), and protocol-level limits (sheet name length, column
  width, row height). Links to Apache POI `SpreadsheetVersion` apidocs and the Microsoft Excel
  specifications page.
- All limit enforcement sites in source code now carry a trailing `// LIM-NNN` comment
  cross-referencing the corresponding registry entry in `docs/LIMITATIONS.md`: `LIM-001` on
  `MAX_WINDOW_CELLS` in both protocol and engine, `LIM-002` on the `.xlsx` path check,
  `LIM-003` on the sheet name length check, `LIM-004` on the column width check, `LIM-005`
  on the row height check, `LIM-006` on the duplicate read request ID check, and `LIM-007` on
  the `GET_CELLS` address list validation.

## [0.11.0] - 2026-03-28

### Added

- `--help` output now includes a one-line product description sourced from `gradle.properties`
  as the single canonical definition.
- Protocol catalog now enumerates five plain (non-polymorphic) request types:
  `CommentInput`, `NamedRangeTarget`, `CellStyleInput`, `CellBorderInput`, and
  `CellBorderSideInput`, with their required and optional fields listed.
- Sheet name length validation: `ENSURE_SHEET`, `RENAME_SHEET`, `SET_NAMED_RANGE`, and every
  other operation that accepts a `sheetName` now rejects names longer than 31 characters with
  `INVALID_REQUEST` before any workbook state is touched.
- Duplicate `requestId` detection in `GridGrindRequest`: construction fails at protocol
  deserialization time when two or more reads share the same `requestId`.
- Email hyperlink address validation: `HyperlinkTarget.Email` now rejects addresses that lack
  `@`, have an empty local part, or have an empty domain with `INVALID_REQUEST`.

### Changed

- Unknown JSON fields in requests are now rejected with `INVALID_REQUEST_SHAPE` instead of
  being silently ignored (`FAIL_ON_UNKNOWN_PROPERTIES` enabled on the protocol mapper).
- Mutation operations (`SET_CELL`, `SET_RANGE`, `APPLY_STYLE`, `SET_HYPERLINK`, `SET_COMMENT`,
  `APPEND_ROW`, `AUTO_SIZE_COLUMNS`) now require the target sheet to exist; they no longer
  auto-create it. Use `ENSURE_SHEET` before the first write to a sheet.
- `GET_CELLS` now returns a blank-typed cell snapshot for addresses that have never been written
  rather than returning `CELL_NOT_FOUND`. Empty cells are valid cells.
- `FORMULA` cell input type description in `--print-protocol-catalog` now notes that the
  leading `=` must be omitted; the engine adds it internally.

## [0.10.0] - 2026-03-28

### Added

- Artifact-emitted discovery commands: `--print-request-template` for a minimal valid request and
  `--print-protocol-catalog` for the machine-readable protocol inventory.
- Protocol-owned catalog metadata covering source types, persistence types, operations, reads, and
  nested tagged request unions.
- Request-shape error classification via the new `INVALID_REQUEST_SHAPE` problem code.

### Changed

- All tagged request unions now use `type` as the discriminator field, including `source`,
  `persistence`, selections, and success-side persistence outcomes.
- CLI help is now first-success capable: it prints the current version, a minimal valid request,
  stdin and Docker-mounted examples, discovery commands, Docker path semantics, and version-routed
  public doc links.
- Failure contexts now surface `sourceType` and `persistenceType` fields, matching the request
  contract terminology.
- Public docs, runnable examples, and committed Jazzer request seeds now use the unified `type`
  contract and describe the new discovery and error-taxonomy behavior.
- Jazzer protocol-request replay and assertions now distinguish syntax failures from request-shape
  failures and semantic validation failures.

### Fixed

- Syntactically valid payloads with missing required fields, unknown discriminator IDs, or wrong
  token shapes are no longer mislabeled as `INVALID_JSON`.

## [0.9.0] - 2026-03-27

### Added

- Finding-bearing document analysis reads: `ANALYZE_FORMULA_HEALTH`,
  `ANALYZE_HYPERLINK_HEALTH`, `ANALYZE_NAMED_RANGE_HEALTH`, and
  `ANALYZE_WORKBOOK_FINDINGS`.
- Real CLI help output via `--help` and `-h`, including usage guidance for stdin/stdout,
  request/response files, Docker, and the packaged Java 26 fat JAR.

### Changed

- Factual read operations now use `GET_*` names consistently: formula surface, sheet schema, and
  named-range surface are exposed as introspection reads instead of conclusion-bearing
  `ANALYZE_*_SURFACE` operations.
- Success persistence outcomes now distinguish `NOT_SAVED`, `SAVED_AS`, and `OVERWRITTEN`, with
  explicit caller-facing path tokens plus execution-environment paths for saved workbooks.
- Internal document-intelligence architecture now uses a shared analysis-finding model across
  formula, hyperlink, named-range, and aggregate workbook health reads.
- Jazzer generators, assertions, readable request seeds, and live-run invariants now enforce the
  renamed read taxonomy and the new persistence/analysis contracts.
- Refreshed pinned GitHub Actions workflow dependencies to current Node 24-ready releases, replaced the release-publish action with a GitHub CLI release step, and configured Dependabot to stop reopening the rejected `gradle/actions` v6 major upgrade.

## [0.8.0] - 2026-03-27

### Changed
- The CLI and protocol architecture now expose a transport-neutral request-executor port and a
  thinner CLI transport boundary, keeping argument parsing, protocol I/O, and execution concerns
  more cleanly separated.
- Public product wording now describes GridGrind as agent-first but not agent-only, matching the
  shipped JSON protocol, CLI/container transport, and current non-MCP distribution model.
- Local Jazzer regression verification now replays committed seeds through four isolated
  per-harness launcher tasks before producing the aggregate regression summary, and local Jazzer
  harness execution no longer depends on Gradle's flaky binary test-results pipeline.
- The request and success protocol now use ordered `reads` and `persistence` outcomes instead of
  the old aggregated `analysis` and nullable saved-path model, making post-mutation workbook
  introspection and insights explicit, typed, and request-correlated.

### Added
- Read operations for workbook summary, named ranges, sheet summary, cells, windows,
  merged regions, hyperlinks, comments, sheet layout, formula surface, sheet schema, and
  named-range surface.
- Public read-heavy example request in `examples/introspection-analysis-request.json`.
- Jazzer request-seed coverage for the new read-heavy public example and the broader `reads`
  protocol shape.

## [0.7.0] - 2026-03-27

### Added
- Cell-metadata authoring with `SET_HYPERLINK`, `CLEAR_HYPERLINK`, `SET_COMMENT`, and
  `CLEAR_COMMENT`, plus typed hyperlink and comment analysis in cell reports and previews.
- Workbook named-range authoring with `SET_NAMED_RANGE` and `DELETE_NAMED_RANGE`, plus requested
  workbook-level named-range analysis and named-range counts in workbook summaries.
- Public example request for hyperlink, comment, and named-range workflows in
  `examples/excel-authoring-essentials-request.json`.
- Jazzer regression seeds covering the public authoring example, hyperlink/comment workflow
  failures, named-range workflow failures, named-range normalization round-trips, and additional
  authoring-metadata replay cases.

### Changed
- Named-range targets are now canonicalized to top-left:`bottom-right` order on input, so shapes
  such as `B2:A1` are stored and reported as `A1:B2`.
- `CLEAR_RANGE` now documents its full effect on cell metadata: it clears hyperlink and comment
  state in addition to values and styles.
- Jazzer coverage and operator docs now describe the expanded authoring surface and the larger
  committed seed floor.

### Fixed
- Repeated `SET_HYPERLINK` writes on the same cell now preserve the latest hyperlink target after
  `.xlsx` save and reopen instead of leaking an older target through the persisted workbook.
- `.xlsx` round-trip verification no longer treats named-range target normalization as a failure
  after save and reopen.

## [0.6.0] - 2026-03-26

### Added
- Local-only Jazzer fuzzing layer in a separate nested Gradle build under `jazzer/`, including
  protocol request fuzzing, structured workflow fuzzing, engine command-sequence fuzzing, and
  `.xlsx` round-trip fuzzing.
- Convenience wrapper scripts under `jazzer/bin/` for regression replay, per-harness fuzzing,
  aggregate fuzzing, and local cleanup.
- Task-specific local corpus storage under `jazzer/.local/runs/*/.cifuzz-corpus/`.
- Jazzer operator commands for latest-summary status/report views, corpus inspection, finding
  listing, one-off input replay, and seed promotion.
- Per-target Jazzer run history with latest-summary JSON/text artifacts, per-harness telemetry,
  and replayed local finding artifacts.
- Committed custom seed floor for all four current Jazzer harnesses, including readable public
  request examples and replay-verified binary workflow seeds, plus promotion metadata under
  `jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata/`.
- Deterministic Jazzer support tests for summary parsing, summary rendering, and `.xlsx`
  round-trip verifier behavior inside the nested `jazzer/` build.
- Richer `.xlsx` style authoring through `APPLY_STYLE`, including `fontName`,
  typed `fontHeight`, `fontColor`, `underline`, `strikeout`, `fillColor`, and side-aware border
  patches with `all` defaults plus top/right/bottom/left overrides.
- Effective style analysis and `.xlsx` round-trip verification for font, fill, and border facts,
  plus a dedicated formatting-depth example request.

### Changed
- Developer documentation now includes an authoritative Jazzer architecture and operations
  reference in `docs/DEVELOPER_JAZZER.md`.
- The Jazzer workflow now uses lock-protected wrapper scripts so only one local Jazzer command
  runs at a time, avoiding concurrent runtime-initialization failures.
- Jazzer documentation is now split into architecture, operations, and coverage references so the
  local fuzzing layer, command surface, and current harness inventory can be read independently.
- `jazzer/bin/fuzz-all` now runs the four active harness scripts sequentially so every harness
  still gets its own lock, run history, summary, and telemetry artifacts.
- Jazzer workflow fuzzing now exercises `NEW` and `EXISTING` source modes together with `NONE`,
  `SAVE_AS`, and `OVERWRITE` persistence modes instead of staying on new-workbook, no-persistence
  flows only.
- Root project-file formatting now excludes local-only instruction and scratch directories, keeping
  local workspace state outside the canonical quality gates.
- `jazzer/bin/list-corpus` now reports generated local corpus and committed custom seeds
  separately, making the active-fuzz seed floor easier to interpret.
- The committed Jazzer regression seed floor now includes the formerly crashing malformed-formula
  engine command sequence as an expected-invalid replay case.
- Public operation and quick-reference docs now describe the expanded style contract, including
  RGB color normalization, solid-fill semantics, border-style enums, and analyzed effective style
  output.
- The public style contract now uses typed `fontHeight` objects instead of the old integer
  `fontSizePoints` field. Requests can express font height as exact points or exact twips, and
  analyzed style reports now return both `fontHeight.points` and `fontHeight.twips`.

### Fixed
- Malformed formulas that trigger Apache POI parser-state `IllegalStateException`s are now
  surfaced as `INVALID_FORMULA` instead of leaking as `INTERNAL_ERROR`.
- Jazzer `.xlsx` round-trip fuzzing now validates persisted formatting-depth style state
  accurately instead of relying on coarse style heuristics, so alignment-only and exact
  `fontHeight` round-trips no longer produce false positives.
- Jazzer latest-summary status and report views now distinguish active findings from expected-
  invalid and replay-clean local artifacts, preventing stale local crash files from being
  misreported as current failures.
- Jazzer summary parsing now handles active-fuzz corpus-size output that uses `Kb` units, keeping
  latest-summary metrics accurate for longer local fuzzing runs.

## [0.5.0] - 2026-03-25

### Added
- Explicit `.xlsx`-only request contract for existing-workbook and save-as paths, with deterministic rejection of `.xls`, `.xlsm`, `.xlsb`, and other non-`.xlsx` workbook paths.
- Reusable `.xlsx` round-trip test infrastructure for reopening generated workbooks and asserting sheet order, merged regions, column widths, row heights, and freeze-pane state.
- Sheet-management operations: `RENAME_SHEET`, `DELETE_SHEET`, and `MOVE_SHEET`.
- Structural layout operations: `MERGE_CELLS`, `UNMERGE_CELLS`, `SET_COLUMN_WIDTH`, `SET_ROW_HEIGHT`, and `FREEZE_PANES`.
- Runnable example requests for sheet management and structural layout workflows.
- Expanded capability documentation describing current Excel support, shipped behavior, and remaining Apache POI parity gaps.

### Changed
- GridGrind workbook support is now an explicit product contract for `XSSF .xlsx` only rather than an inferred capability.
- Public operation and error documentation now covers strict sheet-management and structural-layout semantics, including exact-match unmerge behavior, explicit width and height units, and freeze-pane coordinate rules.
- Developer documentation now includes the expanded runnable example inventory.
- Project documentation now summarizes the shipped `.xlsx`, sheet-management, and structural-layout scope together with verification status and the next recommended parity area.

### Fixed
- Capability inventory contradictions around merged cells and freeze panes removed so the standing agent inventory matches the shipped protocol surface.
- Structural layout coverage gaps in protocol and engine tests closed, including direct helper-branch verification for merge overlap checks, width and height validation, exact merged-region lookup, and freeze-pane coordinate validation.

## [0.4.1] - 2026-03-25

### Added
- `--version` flag: prints `gridgrind <version>` to stdout and exits with code 0.

## [0.4.0] - 2026-03-25

### Security
- All GitHub Actions action references pinned to exact commit SHAs; Dependabot configured to
  keep pins current automatically.

### Changed
- CI workflow now runs `./gradlew check` only; redundant `coverage` task removed from CI since
  report generation without artifact upload produced files that were immediately discarded.
- Job timeouts added to all three workflows: 15 minutes for CI, 20 minutes for Release and
  Container.
- `FormulaException` converted from abstract sealed class to sealed interface; `InvalidFormulaException`
  and `UnsupportedFormulaException` now extend `IllegalArgumentException` directly and implement the
  interface, carrying their own fields.
- `PayloadException` converted from abstract sealed class to sealed interface; `InvalidJsonException`
  and `InvalidRequestException` now extend `IllegalArgumentException` directly and implement the
  interface, carrying their own fields.
- `ExcelRange` converted from a manually managed class to a record with a compact constructor that
  enforces non-negative and ordered bounds.
- `ExcelCellStyleSnapshot` alignment fields changed from `String` to typed `ExcelHorizontalAlignment`
  and `ExcelVerticalAlignment` enums.
- `CellStyleInput` shadow enums `HorizontalAlignmentInput` and `VerticalAlignmentInput` removed;
  fields now use `ExcelHorizontalAlignment` and `ExcelVerticalAlignment` directly from the engine.
- `ExcelPreviewRow` compact constructor added to enforce an unmodifiable defensive copy of the cells
  list and coerce null to an empty list.
- `GridGrindProblems.enrichContext` exhaustively enumerates all nine `ProblemContext` subtypes
  without a `default` arm; compiler now enforces exhaustiveness.
- `GridGrindService.closeWorkbook` converted from `instanceof` conditional to exhaustive pattern-
  matching switch over the sealed `GridGrindResponse` type.
- `ExcelSheet.lastColumnIndex` and `ExcelSheet.save` dead null-check branches removed; POI
  iterators never yield null rows and `Path.toAbsolutePath().getParent()` is always non-null.
- `ExcelRange.parseCell` dead null-check on the address parameter removed.
- `WorkbookStyleRegistry.resolveNumberFormat` extracted as a package-private static method to
  cover the null/blank substitution branch directly without reflection.
- `GridGrindJson.jsonLine` and `GridGrindJson.jsonColumn` widened to package-private to allow
  direct testing of the null-location and non-positive-line-number branches.
- `GridGrindProblems.enrichContext` widened to package-private to allow direct testing of context
  types that are never paired with enrichable exceptions in integration paths.
- `GridGrindService.formulaFor`, `sheetNameFor`, `addressFor`, `rangeFor`, and
  `GridGrindProblems.enrichContext` multi-label pattern case arms split into individual per-subtype
  arms, and `when` guards replaced with if-else blocks inside the relevant arms, eliminating
  JaCoCo false-positive missed branches that Java 26 pattern switches generate for multi-label
  case arms.
- `GridGrindService.validateRequest` `None`/`SaveAs` multi-label arm split into two individual
  arms for the same reason.
- `GridGrindJson.jsonPath` dead `if (index >= 0)` guard removed; `getIndex()` always returns
  a non-negative value for array-position references.
- `ExcelSheet.snapshot` formula expression extracted before the try block so the catch arm can
  reference it without a ternary inside the catch, removing an unreachable branch.
- `GridGrindService.formulaFor` `when op.value() instanceof CellInput.Formula` guard replaced
  with an inner exhaustive switch over the sealed `CellInput` type, eliminating the guard-false
  synthetic branch that JaCoCo counted as a missed branch.

## [0.3.0] - 2026-03-25

### Changed
- `source.mode` wire value renamed: `EXISTING_FILE` is now `EXISTING`.
- `persistence.mode` wire value renamed: `OVERWRITE_SOURCE` is now `OVERWRITE`. The `OVERWRITE`
  mode no longer accepts a `path` field; it always overwrites the source file.
- `AUTO_SIZE_COLUMNS` operation no longer accepts a `columns` field; it always auto-sizes all
  populated columns on the sheet.
- `CellInput.Date` component field renamed from `localDate` to `date`; `CellInput.DateTime`
  component field renamed from `localDateTime` to `dateTime`.

## [0.2.0] - 2026-03-25

### Added
- Native `linux/arm64` container image published alongside `linux/amd64`. Apple Silicon Macs,
  ARM Linux, and Windows ARM pull the correct image automatically with no `--platform` flag.

### Fixed
- Error reference documentation corrected: category values, recovery strategy names, problem
  code table, and causes chain fields now match the actual wire protocol.

## [0.1.0] - 2026-03-24

### Added
- Initial release.

[Unreleased]: https://github.com/resoltico/GridGrind/compare/v0.30.0...HEAD
[0.30.0]: https://github.com/resoltico/GridGrind/compare/v0.29.0...v0.30.0
[0.29.0]: https://github.com/resoltico/GridGrind/compare/v0.28.0...v0.29.0
[0.28.0]: https://github.com/resoltico/GridGrind/compare/v0.27.0...v0.28.0
[0.27.0]: https://github.com/resoltico/GridGrind/compare/v0.26.0...v0.27.0
[0.26.0]: https://github.com/resoltico/GridGrind/compare/v0.25.0...v0.26.0
[0.25.0]: https://github.com/resoltico/GridGrind/compare/v0.24.0...v0.25.0
[0.24.0]: https://github.com/resoltico/GridGrind/compare/v0.23.0...v0.24.0
[0.23.0]: https://github.com/resoltico/GridGrind/compare/v0.22.0...v0.23.0
[0.22.0]: https://github.com/resoltico/GridGrind/compare/v0.21.0...v0.22.0
[0.21.0]: https://github.com/resoltico/GridGrind/compare/v0.20.0...v0.21.0
[0.20.0]: https://github.com/resoltico/GridGrind/compare/v0.19.0...v0.20.0
[0.19.0]: https://github.com/resoltico/GridGrind/compare/v0.18.0...v0.19.0
[0.18.0]: https://github.com/resoltico/GridGrind/compare/v0.17.0...v0.18.0
[0.17.0]: https://github.com/resoltico/GridGrind/compare/v0.16.0...v0.17.0
[0.16.0]: https://github.com/resoltico/GridGrind/compare/v0.15.0...v0.16.0
[0.15.0]: https://github.com/resoltico/GridGrind/compare/v0.14.0...v0.15.0
[0.14.0]: https://github.com/resoltico/GridGrind/compare/v0.13.0...v0.14.0
[0.13.0]: https://github.com/resoltico/GridGrind/compare/v0.12.0...v0.13.0
[0.12.0]: https://github.com/resoltico/GridGrind/compare/v0.11.0...v0.12.0
[0.11.0]: https://github.com/resoltico/GridGrind/compare/v0.10.0...v0.11.0
[0.10.0]: https://github.com/resoltico/GridGrind/compare/v0.9.0...v0.10.0
[0.9.0]: https://github.com/resoltico/GridGrind/compare/v0.8.0...v0.9.0
[0.8.0]: https://github.com/resoltico/GridGrind/compare/v0.7.0...v0.8.0
[0.7.0]: https://github.com/resoltico/GridGrind/compare/v0.6.0...v0.7.0
[0.6.0]: https://github.com/resoltico/GridGrind/compare/v0.5.0...v0.6.0
[0.5.0]: https://github.com/resoltico/GridGrind/compare/v0.4.1...v0.5.0
[0.4.1]: https://github.com/resoltico/GridGrind/compare/v0.4.0...v0.4.1
[0.4.0]: https://github.com/resoltico/GridGrind/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/resoltico/GridGrind/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/resoltico/GridGrind/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/resoltico/GridGrind/releases/tag/v0.1.0
