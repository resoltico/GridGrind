# Changelog

Notable changes to this project are documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[Unreleased]: https://github.com/resoltico/GridGrind/compare/v0.7.0...HEAD
[0.7.0]: https://github.com/resoltico/GridGrind/compare/v0.6.0...v0.7.0
[0.6.0]: https://github.com/resoltico/GridGrind/compare/v0.5.0...v0.6.0
[0.5.0]: https://github.com/resoltico/GridGrind/compare/v0.4.1...v0.5.0
[0.4.1]: https://github.com/resoltico/GridGrind/compare/v0.4.0...v0.4.1
[0.4.0]: https://github.com/resoltico/GridGrind/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/resoltico/GridGrind/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/resoltico/GridGrind/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/resoltico/GridGrind/releases/tag/v0.1.0
