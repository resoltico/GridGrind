# Changelog

Notable changes to this project are documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- Extended the structural-governance stack beyond handwritten Java: root `check` now runs
  `verifyControlPlaneShape`, which scans repository-owned shell gates, Kotlin build-logic, the release
  protocol, and the public changelog ledger against expiring reviewed budgets from
  `gradle/control-plane-shape-policy.tsv` so the repo-governing control plane cannot drift outside the same
  no-god-file ratchet model as product code.
- Tightened the forbidden tagged-union and god-record enforcement so the build now inspects named nested
  record/class variants as well as direct top-level declarations, closing the blind spot around the
  sealed-interface-with-nested-record style GridGrind uses for most domain surfaces.
- Split the packaged CLI discovery and help surface across narrower role-owned seams: identity commands,
  example/task discovery, protocol-catalog output, trailing-argument validation, and help-section
  rendering now live on dedicated helpers instead of one broad catalog-command implementation.

### Fixed

- Finished the OOXML encryption, custom-XML, and grouped table-report hard-break migration across the
  remaining Jazzer and parity verification surfaces, including invariant checks, promoted test fixtures, and
  Stage 3 support tests, so the regression and replay stack now validates the live sealed DTO model instead
  of stale flat constructors.
- Aligned invalid-request-shape reporting across the CLI doctor, runtime problem surface, and JSON codec
  layer: missing required root fields and explicit `null` placeholders now classify as `INVALID_REQUEST_SHAPE`,
  and persist-workbook collision reporting no longer dereferences absent save-as paths while building public diagnostics.
- Hardened deterministic OOXML artifact persistence: the package-copy helper now validates its target path
  explicitly, and deterministic ZIP-package rewrites now clean their temporary output artifact before rethrowing
  any rewrite failure.
  
## [0.68.0] - 2026-06-12

### Changed

- Refreshed the shared dependency baseline to current stable releases where verified newer
  versions are available: Jackson databind `3.2.0`, Error Prone `2.50.0`, NullAway `0.13.6`,
  PMD `7.25.0`, and JaCoCo `0.8.15`.
- Updated the GitHub Actions pins to `actions/checkout` `6.0.3` and `gradle/actions` `6.1.1`,
  including the Gradle wrapper-validation workflow.
- Completed the compact cell-payload model migration across the contract, examples, and tooling:
  structured row and range writes now revolve around the typed `CellRowInput`, `CellGridInput`,
  and `CellScalarValue` families end to end, while selector validation and JSON-problem
  translation now live on narrower role-owned seams instead of monolithic helper buckets.
- Re-segmented the developer-facing authoring and discovery internals around role-owned helpers:
  Java authoring now teaches `ExpectedValues` plus grouped workbook/sheet/inspection query
  families, and the CLI/runtime discovery stack now derives its public catalog, keyword-match,
  task-plan, and failure-report surfaces from dedicated parser/index/rendering seams rather than
  broad utility hubs.
- Tightened the structural-governance stack again: repo-owned source-shape budgets now scan every
  handwritten Java source set instead of only `src/main/java`, the dedicated semantic-shape PMD
  gate now owns `GodClass` and `CouplingBetweenObjects` with explicit reviewed exceptions, the
  source-shape family ratchet now covers field and nested-type budgets plus stale prefix headroom,
  and the duplication gate now trips earlier on materially duplicated Java blocks.
- Expanded the packaged discovery-execution verifier so it can execute the installed `gridgrind`
  launcher script directly in addition to the fat JAR and Docker image, keeping the real
  end-user binary surface under the same example and task-starter field-test discipline.
- Refreshed the operator and developer docs around the live runtime contract: the root README and
  quick references now distinguish the compact protocol index from full catalog dumps, teach the
  packaged launcher alongside the JAR and Docker entry points, document `SUMMARY` as the default
  journal level, and describe the current semantic-shape/build-logic ownership model without
  carrying forward the older release assumptions.

### Fixed

- Removed the JaCoCo prerelease-only repository wiring and timestamped snapshot coordinate now
  that Java 26 support is available on JaCoCo `0.8.15` GA.
- Removed the stale Dependabot major-version suppression for `gradle/actions`, so repository
  automation no longer hides future workflow-runtime upgrades.
- Closed the remaining structural-governance blind spots around the new god-file stack:
  interface methods now count toward public-surface budgets, reviewed source-shape and
  semantic-shape waivers are re-evaluated against live UTC dates and build-logic freshness tests,
  repo-wide verification now rejects any reintroduced legacy `buildSrc` tree, and the split
  hyperlink file-normalization seam is covered end to end so the contract coverage gate stays
  truthful.
- Realigned the packaged CLI verification and smoke stack with the live artifact contract:
  bare no-request failures are now verified on stderr, compact and full protocol-catalog surfaces
  are checked separately, request-template verification expects the shipped `journal.level=SUMMARY`
  default, and Docker publication fakes mirror the same stderr/catalog behavior instead of
  preserving the older release-surface shape.
- Repaired the Docker `STREAMING_WRITE` smoke scenario and selector-contract guard to use the
  current `CellRowInput` wrapper for `APPEND_ROW`, preventing stale flat row payloads from hiding
  in release-surface shell requests until runtime.
- Brought Jazzer support generation and promoted protocol-request seeds back into line with the
  compact cell-payload contract by treating authored error cells as value-bearing during XLSX
  round-trip expectations and refreshing the promoted request metadata after `PrintSetupInput`
  became a fully required shape.

## [0.67.0] - 2026-06-05

### Changed

- Hardened the repository-wide structural-governance contract beyond file-size ratcheting. Root
  `check` now runs both `verifyJavaSourceShape` and `verifyJavaSourceDuplication`, the shared
  `gradle/source-shape-policy.tsv` now owns `duplicationGuard` plus per-surface
  `reviewExpiresOn` and `splitTrigger` metadata, reviewed exact-surface overrides now fail when
  they drift too far above the current file shape, and the included `gradle/build-logic` tests
  are wired into the root verification path through configuration-cache-safe compile-output
  cleaning so the governance owner cannot drift separately from the product build.
- Consolidated the open post-release dependency queue into one current-`main` maintenance line:
  Kotlin Gradle plugin `2.4.0`, Jackson databind `3.1.4`, Jackson annotations `2.22`, Shadow
  `9.4.2`, Spotless `8.6.0`, NullAway `0.13.5`, JaCoCo trunk build `0.8.15.202606040741`
  (published as `0.8.15-20260604.194125-118`), Jakarta XML Bind API `4.0.5`, and the pinned
  `azul/zulu-openjdk-alpine:26-jre` runtime digest used by the published Docker image.
- Refreshed the GitHub Actions publication and CI tooling pins to the current Docker-maintained
  releases: `docker/setup-buildx-action` `4.1.0`, `docker/setup-qemu-action` `4.1.0`,
  `docker/login-action` `4.2.0`, `docker/metadata-action` `6.1.0`, and
  `docker/build-push-action` `7.2.0`.
- Clarified the release protocol's Dependabot hygiene step so stale, post-release maintenance
  heads can be replaced by one deliberate current-`main` maintenance PR instead of being merged
  through repeated branch-refresh churn one by one.
- Clarified the release protocol's red-`Gate` handling for hosted dependency-resolution flake:
  before rewriting build logic or pinned versions around an external-repository `403`/`404`/`5xx`
  failure, the release flow now requires a cold local `--refresh-dependencies` repro and one
  explicit `gh run rerun --failed` pass so transient infrastructure noise is separated from real
  release-branch breakage.
- Tightened the execution contract around explicit invocation-owned state: the public Java API now
  requires both request-root and temp-root inputs, stdin-driven CLI execution/doctoring now
  requires `--execution-root <path>` instead of inheriting ambient process cwd semantics, the
  engine no longer falls back to `~/.gridgrind/tmp`, and the CLI/docs/help surface now teaches
  `.gridgrind/tmp` under the request root or explicit execution root as the default scratch
  location unless `--temp-root <path>` overrides it.
- Refined the public CLI discovery and help contracts around the verified artifact surface:
  `--print-protocol-catalog --search` now returns compact summary-first matches instead of inline
  full-entry payload dumps, `--response` write failures now fall back to the originating command
  family’s own machine-readable failure shape instead of masquerading as execution responses, the
  split help surfaces now render command invocations as commands rather than prose-wrapped tables,
  and the jar/Docker CLI contract verifier now parses the first product-owned JSON document from
  PTY output instead of treating terminal wrapping or runtime trailers as product drift.
- Tightened the root `README.md` front-door copy so it matches the current artifact contract more
  precisely: it now distinguishes workbook-save atomicity from response-file output, teaches the
  explicit `--execution-root <path>` rule for stdin-driven first-contact flows, and no longer
  describes asset-backed example directories as universally blank-workspace runnable.

### Fixed

- Made the packaged discovery-execution verifier emit per-step liveness heartbeats while waiting
  on slow example or task-starter runs, so long-running release-surface checks no longer go quiet
  long enough to trip `check.sh`'s watchdog even when the underlying CLI execution is healthy.

## [0.66.0] - 2026-05-27

### Changed

- Replaced the contract-test-only source-size watchlist with a repository-wide Java source-shape
  quality gate wired into root `check`: `verifyJavaSourceShape` now parses every production Java
  source, writes `build/reports/source-shape/source-shape.tsv`, and enforces role-owned budgets
  from `gradle/source-shape-policy.tsv`, while `ArchitectureSeamAuditTest` now focuses on
  workbook seams, module exports, and documentation parity instead of carrying the size ratchet
  itself.
- Hard-split the workbook authoring API into grouped role surfaces. Workbook creation/opening now
  lives on `ExcelWorkbooks`; `ExcelWorkbook` now exposes dedicated `formulas()`, `sheets()`,
  `customXml()`, `protection()`, `names()`, `tables()`, `pivots()`, and `persistence()` surfaces;
  `ExcelSheet` now exposes grouped `cells()`, `annotations()`, `drawings()`, `metadata()`,
  `layout()`, `rows()`, and `columns()` surfaces instead of one catch-all public API.
- Decomposed broad engine and verification internals into focused helpers: conditional-formatting
  style IO now separates write, snapshot, color, and border-style concerns; sheet-copy replay now
  separates named-range, validation, conditional-formatting, autofilter, and table
  responsibilities; OOXML package security now separates orchestration, encryption, inspection,
  signing, and file persistence; execution workflow and source-backed resolution now split into
  dedicated direct-event, streaming, chart-input, mutation-action, and structured-input helpers;
  Jazzer analysis, cell, and fuzz-data support now each own narrower verification slices instead
  of carrying monolithic support blobs.
- Applied the semantic-shape PMD profile to `contract`, `engine`, and `excel-foundation`
  `pmdMain`, and corrected the developer docs to describe the real PMD, source-shape, and
  downstream-coverage ownership model used by the build.
- Updated the shared build/test dependency baseline to JUnit `6.1.0`, JaCoCo trunk build
  `0.8.15.202606030734` (published as `0.8.15-20260603.073432-117`), Log4j `2.26.0`,
  SLF4J `2.0.18`, and Jakarta Activation `2.1.4`.
- Refreshed the root `README.md` to match the current CLI first-run surface: it now shows the
  repository JAR path, the local Docker build path, the built-in example and doctor flow, the
  current discovery commands, and the split between `--help`, `--help-protocol`, and
  `--help-guidance`.
- Aligned the packaged distribution identity with the executable product name: the shadow install
  tree now materializes under `build/install/gridgrind`, and the packaged archives now publish as
  `gridgrind-<version>.zip` and `gridgrind-<version>.tar` instead of `cli-shadow-*`.
- Changed the repository-root runtime Docker image build into a self-contained multi-stage build:
  `Dockerfile` now compiles `:cli:shadowJar` inside a pinned builder stage, `docker-smoke.sh`
  verifies that path directly, and the user-facing quick-start/help/docs now teach the local
  `docker buildx build --load -t gridgrind-local .` flow alongside the published GHCR image.
- Rebuilt the CLI discovery surface around typed public contracts instead of stitched protocol
  sketches: built-in examples and official task starters now publish `workspaceMode` plus
  `requiredPaths`, `--print-task-plan` emits task-owned executable starter requests rather than
  generic step-template composites, and the discovery/help renderers now present catalogs,
  workflow guidance, operator notes, and follow-on commands as separate structured sections.
- Tightened the CLI’s first-contact recovery and intent guidance: `--doctor-request` now performs
  request-tree preflight and batches every independently provable blocking problem into one doctor
  report, argument failures now return nearest-flag suggestions plus command-family-specific
  recovery text, and task keyword discovery now withholds low-confidence candidates instead of
  over-recommending workflows on nonsense queries.
- Expanded the release-surface verification contract from help-only checks to black-box execution:
  the jar and Docker verification stack now runs every published built-in example and official
  task starter, verifies the packaged discovery/help surface directly, and keeps the no-stdin CLI
  contract deterministic even when the parent shell is interactive.

### Fixed

- Replaced destructive `build/libs` cleanup with per-project stale-jar pruning, so composite,
  root, and nested verification runs no longer lose sibling module jars during `check`,
  `shadowJar`, or Jazzer verification.
- Removed the legacy `EXPECT_PRESENT` / `EXPECT_ABSENT` compatibility hint path from JSON
  request-shape failures. Deleted assertion names are now rejected as unknown types without
  fallback suggestion text for retired vocabulary.
- Removed POI-private hyperlink field/constructor dependence from sheet-clone preparation by
  rehydrating external hyperlinks through a materialized relation wrapper, and repaired blank
  external relationship ids before clone-time relation creation so copied sheets reopen with
  canonical hyperlink relationships.
- Kept the nested Jazzer coverage subset aligned with its intended deterministic support-contract
  surface after the fuzz-decoder split by excluding `FuzzAddressDecoders`, `FuzzStyleDecoders`,
  and `FuzzValueDecoders` alongside `FuzzDataDecoders`.
- Added machine-readable JSON mode to the local Jazzer operator wrappers for `status`, `report`,
  `list-findings`, and `list-corpus`, and hardened the wrapper regression suite so the help text,
  JSON passthrough, and lock-conflict wording remain wrapper-owned instead of leaking Gradle task
  internals.
- Moved the contract JSON tests onto Jackson 3 context-aware parser entry points and replaced raw
  reflective coverage casts in the engine helper tests with typed extraction helpers, removing the
  deprecated-parser and unchecked-cast warning churn from the release gates.
- Made the published-example and task-starter discovery verifier emit per-plan progress heartbeats
  during both jar and Docker release-surface checks, and added a focused shell regression so the
  merge-handoff gate cannot fail on quiet watchdog timeouts while black-box artifact verification
  is actively running.

## [0.65.0] - 2026-05-15

### Added

- Decomposed `GridGrindProtocolCatalogLookupSupport` (936 lines, 50+ methods) into four
  package-private helpers: `CatalogSearchRankingSupport` (rank scoring, `isTopLevelPublishedGroup`,
  `stepTemplateFor`), `CatalogShapeTextSupport` (OOXML shape text rendering and shape-reference
  traversal), `CatalogSearchAggregationSupport` (match grouping, dedup, `SearchAggregate`), and
  `CatalogRefResolutionSupport` (flat catalog ref building); `GridGrindProtocolCatalogLookupSupport`
  retains only the five public entry points, search coordination, value-type hierarchy, and the
  `CatalogLookupRef` and `RankedSearchMatch` shared records.

- Added `EXPECT_SHEET_PRESENT` and `EXPECT_SHEET_ABSENT` assertion types to `PresenceAssertion`,
  closing the sheet-presence gap that existed alongside `EXPECT_NAMED_RANGE_PRESENT`,
  `EXPECT_TABLE_PRESENT`, `EXPECT_PIVOT_TABLE_PRESENT`, and `EXPECT_CHART_PRESENT`; both types
  accept all `SheetSelector` variants (`SHEET_ALL`, `SHEET_BY_NAME`, `SHEET_BY_NAMES`) and
  evaluate by matching against the live workbook sheet list without a separate inspection query;
  added `WorkbookInspectionResult.SheetsResult` as the corresponding observation carrier and wired
  the two new arms into `AssertionObservationExecutor` and `AssertionExecutor`.


- Added `passed` field to `AssertionResult` so every assertion outcome carries an explicit
  boolean: `true` when the assertion held, `false` when it failed. Failure responses now include
  the partial assertion list accumulated before the first failure, making the full per-step
  assertion picture available without cross-referencing journal step events.
- Added `assertions` field to `GridGrindResponse.Failure` so assertion steps executed before the
  first failure are visible in the failure response alongside the `problem` payload.
- Added `protocolVersion` field to `CatalogSearchResult` so `--print-protocol-catalog --search`
  responses carry the same top-level version identifier as the full catalog response.
- Added dedicated `CliFailureReport` shape for CLI argument, lookup, and command-usage errors,
  emitting a machine-readable JSON payload with `exitCode`, `command`, `code`, `message`,
  `location`, `argument`, `suggestions`, and `resolution` instead of a bare text message.

### Changed

- Bumped Gradle from `9.5.0` to `9.5.1`.
- Replaced the 6 intermediate package-private constructors in `GridGrindCli` with two static
  `forTesting(executor)` and `forTesting(executor, standardInputIsInteractive)` factory methods;
  the public no-arg and canonical 6-arg constructors are unchanged; deleted four tests that
  verified constructor argument counts by name rather than observable behavior and updated all
  7 affected test call sites.
- Replaced the hardcoded `"source.type must be NEW; mutation actions limited to ENSURE_SHEET"`
  prose snippet in `verify-cli-contract.sh` with a catalog-derived check: `ENSURE_SHEET` and
  `APPEND_ROW` are now verified against `protocol_help_output` immediately after the existing
  catalog STREAMING_WRITE structural checks, so the canonical constraint terms drive both
  surfaces instead of the shell script carrying a parallel string literal.

- Moved top-level catalog category enumeration into `Catalog.topLevelGroups()` so the canonical
  list of bare-name-addressable categories (`sourceTypes`, `persistenceTypes`, `stepTypes`,
  `mutationActionTypes`, `assertionTypes`, `inspectionQueryTypes`) is co-located with the record
  fields it enumerates; `GridGrindProtocolCatalogLookupSupport` now derives the lookup refs from
  that method instead of maintaining a separate parallel list.
- Made `EXPECT_PRESENT` and `EXPECT_ABSENT` unknown-type error suggestions self-healing: the
  message now derives the candidate list from the live assertion type registry filtered by
  `_PRESENT` / `_ABSENT` suffix instead of a hardcoded string; `EXPECT_ANALYSIS_FINDING_PRESENT`
  and `EXPECT_ANALYSIS_FINDING_ABSENT` are now included in the suggestions, and the candidate list
  will automatically include any new presence-assertion families added in future releases.


- Replaced `AssertionResult.passed` (`boolean`) with `AssertionResult.outcome`
  (`AssertionOutcome` enum, values `PASSED` / `FAILED`): the boolean emitted `"passed": true` or
  `"passed": false` in JSON, which was inconsistent with the rest of the protocol vocabulary where
  outcomes and statuses use uppercase string enums; the serialised field key is now `"outcome"`
  with values `"PASSED"` and `"FAILED"`.
- Changed `GridGrindResponse` top-level status discriminator values from `SUCCESS` / `ERROR` to
  `SUCCEEDED` / `FAILED`, aligning the response envelope vocabulary with `ExecutionJournal.Status`
  (`SUCCEEDED` / `FAILED`) and `ExecutionJournal.StepOutcome` (`SUCCEEDED` / `FAILED`); the
  response Java types (`GridGrindResponse.Success`, `GridGrindResponse.Failure`) are unchanged.

- Added `LimitationsTraceabilityTest`, a build-time JUnit 5 test that validates bidirectional
  LIM-NNN consistency: every `// LIM-NNN` comment in Java source must have a matching
  `### LIM-NNN` entry in `LIMITATIONS.md` (no orphaned markers), and every `LIMITATIONS.md` Code
  field that references `// LIM-NNN` must have at least one matching comment in Java source (no
  undocumented enforcement sites); the test runs as part of `:contract:check` and receives the
  project root via a Gradle system property. The source-directory scan now auto-discovers all
  `src/main/java` directories under the project root (excluding `/build/` paths) so new modules
  are covered automatically without editing the test. The heading parser now uses
  `LIM_HEADING_PATTERN` (`^### (LIM-[A-Z0-9]+)`) regex consistent with `LIM_PATTERN`, eliminating
  the substring crash that would occur if a LIMITATIONS.md entry heading had no trailing title
  text.

### Fixed

- Added `Objects.requireNonNull(cell)` to the 1-arg `ExcelChartSourceSupport.scalarFromFormula`
  overload, matching the guard already present in the 2-arg overload; the inconsistency meant the
  contract differed silently between two methods with the same name. Extracted a `row` local in
  `resolveChartSource` to replace a double `getRow()` call in the ternary with a single read,
  removing the stylistic pattern that mirrors the `scalarText` null-dereference.

- Fixed a null-dereference in `ExcelChartSourceSupport.scalarText()`: the method chained
  `.getRow()` directly onto `workbook.getSheet(sheetName)` without first checking whether the
  sheet exists; `Workbook.getSheet(name)` returns null for an unknown sheet name, and throws
  `NullPointerException` when `CellReference.getSheetName()` is null (unqualified reference).
  The fix extracts `targetSheet` using the same `requireSheet()` pattern already used by every
  other `getSheet` call site in the class: null sheet name (unqualified reference) falls back to
  the context sheet; a non-null name is validated through `requireSheet`, which throws
  `IllegalArgumentException` with a clear message if the sheet is absent.


- Updated `LIMITATIONS.md` Code fields for LIM-006A and LIM-018 to include `// LIM-NNN`
  traceability notation, aligning the registry entries with the enforcement-site comments added
  to `WorkbookStepValidation.requireStepId` and all eight `rejectDestructiveNamed*` methods
  across `ExcelRowColumnStructureController` and `ExcelRowColumnStructureGuardSupport`.


- Narrowed the `NullPointerException` arm of `GridGrindJsonProblemMessageSupport.validationCause()` so
  that only explicit null-check NPEs (those whose message ends with `"must not be null"`) are
  classified as user-input validation errors; JVM-generated NPEs and NPEs with blank or null
  messages now fall through to `InvalidJsonException` instead of leaking an internal pointer into
  the user-facing error message.
- Updated synthetic `MismatchedInputException` floating-point messages in `GridGrindJsonCoverageTest`
  and `GridGrindJsonTest` from the Jackson 2.x-style `"Floating-point value (X) out of range of
  int"` prefix to the actual Jackson 3.x format `"Cannot coerce Floating-point value (X) to \`int\`
  value (but could if coercion was enabled using \`CoercionConfig\`)"`, keeping the unit-test
  fixtures aligned with real Jackson 3.x exception messages; updated the `cleanJacksonMessage`
  fixture in `GridGrindJsonTest` from the Jackson 2.x `(set X to allow)` hint to the Jackson 3.x
  `(but could if coercion was enabled using \`CoercionConfig\`)` hint, confirming the replacement
  regex strips correctly.
- Replaced the stale `(set [^)]*to allow)` Jackson 2.x coercion-hint removal regex in
  `GridGrindJsonProblemMessageSupport.cleanJacksonMessage()` with the Jackson 3.x equivalent
  `(but could if coercion[^)]*\)` so coercion hints are stripped from user-facing messages on
  the current Jackson version.
- Removed the `"Cannot map \`null\` into type"` branch from `mismatchedInputMessage()` and its
  helper `nullIntoPrimitiveMessage()`: the `GridGrindJsonCodecSupport.rejectExplicitNullMembers()`
  pre-pass rejects all explicit JSON nulls with the full dotted path before Jackson deserialization
  starts, making the `MismatchedInputException` null-coercion path unreachable via any
  `GridGrindJson` read method; removed the synthetic-exception tests that were exercising this
  dead code, and added `rejectsExplicitNullInDeepNestedRequestFieldWithFullDottedPath` to verify
  the actual pre-pass behavior for deeply nested fields.

- Changed `GridGrindJsonCodecSupport.requireNonNullRoot()` message from
  `"problem: <root> must not be null"` to `"JSON payload must not be null"`: the old wording
  carried the legacy `"problem: "` prefix that was removed from all other null-rejection paths
  when `rejectExplicitNullMembers` was refactored to throw `InvalidRequestException` directly;
  updated the two byte-array and stream overload assertions in
  `rejectsTopLevelAndArrayNullRequestPayloads` in `GridGrindJsonCoverageTest` to match.
- Removed the `(?:problem: )?` optional prefix from `NULL_FIELD_PROBLEM_PATTERN` in
  `GridGrindJsonProblemMessageSupport`: the only code that ever emitted the `"problem: "` prefix was
  `rejectExplicitNullMembers`, which now throws `InvalidRequestException` directly; the pattern's
  live production path is `Objects.requireNonNull` NPEs from `@JsonCreator` constructors, whose
  messages have no prefix; updated the two synthetic test cases in
  `helperMethodsCoverFallbackAndLocationEdgeCases` from `IllegalArgumentException("problem: X must
  not be null")` to `NullPointerException("X must not be null")` to reflect the actual production
  trigger format.
- Unified the JSON path-rendering algorithm under `renderPath(List<JacksonException.Reference>)`:
  `jsonPath(JacksonException)` now delegates to `renderPath` instead of containing an identical
  loop; `floatingPointIntoIntegerMessage` already called `renderPath` directly; the canonical
  algorithm now lives in one place with one set of tests.

- `rejectExplicitNullMembers()` now throws `InvalidRequestException` directly with
  `Optional.of(childPath)` as the structured `jsonPath()` instead of a plain
  `IllegalArgumentException` carrying the path in the message text; `decodeValue()` no longer
  needs a catch-and-wrap for that case, so `jsonLocation().jsonPath()` is now correctly populated
  for all explicit-null rejections (previously it was always `Optional.empty()`).
- `floatingPointIntoIntegerMessage()` now renders the full Jackson path when the terminal path
  reference is an array index rather than a named property; the previous implementation returned
  the generic `"JSON value must be an integer value"` in that case even though the structured path
  was available; the new wording is `"JSON value at '<path>' must be an integer value"`.

- Bumped JaCoCo from `0.8.15-20260508.112122-100` to snapshot build `0.8.15.202605130743`
  (2026/05/13), published Maven coordinate `0.8.15-20260513.074320-106`.
- Bumped Kotlin from `2.4.0-Beta2` to `2.4.0-RC`.
- Replaced payload JSON cursor storage on contract payload exceptions with one typed
  `PayloadLocation` model and rewired engine-side read-request problem enrichment to consume that
  normalized location variant directly instead of rebuilding path/line/column state from three
  separate optional values.
- Rebuilt CLI task discovery around typed `TaskDiscoveryProfile` and `TaskIntentProfile` value
  objects, split the old reporting/workbook task bag files into per-workflow definition modules,
  and reweighted English keyword search toward typed goals and artifact kinds instead of incidental
  prose matches.
- Reworked protocol-catalog search into grouped top-level results with attached `stepTemplate`,
  `relatedEntryIds`, and `supportingMatches`, and derived nested/plain field-shape metadata from
  the canonical descriptor groups instead of four parallel registry maps.
- Split the highest-churn engine workbook internals into dedicated `event`, `stream`, `pivot`,
  `customxml`, `ooxml`, `drawing`, and `validation` subpackages so the workbook runtime no longer
  depends on one flat `dev.erst.gridgrind.excel` catch-all namespace.


- Routed all structured JSON output (execution failures, CLI argument failures, doctor-report
  failures) to stdout regardless of exit code. Previously, non-zero responses were written to
  stderr when no `--response` path was configured, breaking piped consumers such as `jq`.
- Routed the bare-invocation no-request failure to stderr instead of stdout: when `gridgrind`
  is called with no `--request` argument and standard input is not a pipe, the
  `INVALID_ARGUMENTS` failure report is now written to stderr (and stdout is left empty);
  callers that supply `--response <path>` still receive the failure in the response file with a
  stderr pointer line; this preserves stdout for downstream consumers in pipelines while keeping
  the failure report visible to interactive operators.
- Replaced deprecated `JsonNodeCreator.textNode(String)` calls in `CatalogStepTemplateSupport`
  with `StringNode.valueOf(String)`, eliminating the Jackson 3.x deprecation warning emitted
  on every compile.
- Consolidated CLI failure stream-routing contract into paired `cliFailureOnStdout` and
  `cliFailureOnStderr` test helpers: `cliFailureOnStdout` asserts stderr is empty and reads
  from stdout (used for `--request`-path failures where stdout is the response channel);
  `cliFailureOnStderr` asserts stdout is empty and reads from stderr (used for bare-invocation
  tests where the no-request failure is routed to stderr); replaces per-test ad-hoc byte-array
  reads that broke silently when routing changed.
- Added a `everyLiveCatalogStepTypeHasAnExecutableTemplate` contract test to catch any future
  omission when a new step type is added to the live catalog without a synthesized template.
- Clarified `CliResponseWriter.writeCliFailureReport` Javadoc to document that the `stderr`
  parameter is used only in the response-file path, not when writing directly to stdout.
- Added `WorkbookInspectionResult.SheetsResult` to the exhaustive switch statements in
  `WorkbookInvariantResponseChecks` and `WorkbookInvariantInspectionResultChecks`; the missing
  arm caused a compile error in the Jazzer module after `SheetsResult` was added for the
  `EXPECT_SHEET_PRESENT` / `EXPECT_SHEET_ABSENT` assertions.
- Updated `docker-smoke.sh` status checks from `"SUCCESS"` to `"SUCCEEDED"` to match the
  `GridGrindResponse` envelope vocabulary change; the five grep assertions now consistently
  use the protocol-aligned discriminator value.
- Added `/.gradle-home/` to `.gitignore` and to the hygiene-script local-state and
  generated-state allowlists; the project-local Gradle user home was appearing as an
  unexpected root entry when `GRADLE_USER_HOME` pointed at the project root.
- Removed a duplicate "REQUIRES_EXAMPLE_ASSETS needs copied examples/ assets beside the request"
  sentence from the `--help-guidance` Discovery section; the same information was already
  expressed in the preceding workspace-mode description line.
- Replaced the vague "Use bare group names first and qualify only when ids collide" phrasing in
  the `--lookup` flag description and the protocol-catalog Discovery notes with concrete examples
  of what is addressable by bare name: individual type ids (SET_CELL, ENSURE_SHEET, GET_CELLS),
  nested/plain type-group names (cellInputTypes, calculationStrategyTypes), and top-level
  operation category names (mutationActionTypes, assertionTypes, inspectionQueryTypes) all
  resolve unqualified.
- Added `--print-protocol-catalog --search <term>` as the first suggestion and updated the
  resolution text for `INVALID_REQUEST_SHAPE` failures so agents and operators see the right
  discovery command when a type discriminator value is wrong; the previous guidance pointed only
  to `--doctor-request` and `--help-protocol`, neither of which lists valid type values.
- Added `// LIM-001` traceability comments to `SelectorValueValidation.requireWindowSize` and
  `WorkbookReadCommand.requireWindowSize`, closing the audit gap identified in the limitations
  registry for the 250,000-cell read-window cap; also corrected the `LIMITATIONS.md` Code field
  which incorrectly named `requireWindowWithinLimit` instead of `requireWindowSize`.
- Added `// LIM-006A` traceability comment to `WorkbookStepValidation.requireStepId` to mark
  the step-id character-set enforcement site for the LIM-006A registry entry.
- Added `// LIM-018` traceability comments to all eight `rejectDestructiveNamed*` methods
  across `ExcelRowColumnStructureController` and `ExcelRowColumnStructureGuardSupport` to mark
  the structural-edit / named-range guard sites for the LIM-018 registry entry.
- Extended `--lookup` to accept top-level operation category names (`mutationActionTypes`,
  `assertionTypes`, `inspectionQueryTypes`, `sourceTypes`, `persistenceTypes`, `stepTypes`);
  previously, bare category names produced a "no match found" error and users had to read the
  full catalog to see entries for a given category.
- Enhanced the `INVALID_REQUEST_SHAPE` / unknown type-discriminator error message to include
  similar valid type-id candidates when the typed value is close (within edit distance) to one
  or more valid subtypes of the target sealed interface; for example, typing `SET_CEL` in an
  action field now produces `Unknown type value 'SET_CEL'; similar valid values: SET_CELL`.


- Corrected `requestType.optionalFields` in the protocol catalog: only `planId` is genuinely
  optional at runtime. The previous catalog listed `protocolVersion`, `persistence`, `execution`,
  `formulaEnvironment`, and `steps` as optional, which contradicted the actual enforcement.
- Replaced placeholder starter filenames and protocol-catalog template values with concrete
  sample data, so the public starter surfaces no longer publish unfinished vocabulary as part of
  the executable contract.
- Tightened the public help and release-surface contract together: protocol help now repeats the
  `STREAMING_WRITE` source/mutation rule, the array-formula rejection rule, and the `do not send
  step.type` rule as standalone contract lines; guidance help now repeats the asset-backed example
  portability rule as one explicit sentence; and the CLI-contract verifier now reads the shipped
  example-catalog `suggestedRequestPath` field instead of the removed `fileName` field.
- Completed the short `--help` synopsis so its `Usage` and `Primary Commands` sections now list
  every first-contact command surface, including example printing, task planning, keyword task
  search, protocol search and lookup, and the `--help`, `--version`, and `--license` info paths.
- Corrected CLI conflict diagnostics so `--help-protocol`, `--help-guidance`, and `-h` report
  their exact authored command token instead of collapsing all help-surface conflicts to
  `--help`.
- Cleaned stale packaged CLI distribution archives out of `cli/build/distributions` before new
  `distZip`, `distTar`, `shadowDistZip`, and `shadowDistTar` outputs are produced, eliminating
  leftover older-version archives after fresh packaging runs.
- Expanded `INPUT_SOURCE_NOT_FOUND` recovery guidance to explain that relative authored-input paths
  root from the request file directory when the CLI reads a request through `--request <path>`.
- Synchronized the post-`0.64.0` operator and developer docs with the shipped CLI surface: example
  printing now documents `--print-example --lookup <id>`, task planning documents
  `--print-task-plan --lookup <id>`, keyword search documents
  `--print-task-keyword-match --query <text>`, bare invocation documents the non-zero no-request
  failure path, compact CLI failure reports are documented as distinct from execution journals, and
  example portability plus exact `examples/...` asset paths are described consistently across the
  quick-start, examples, error, operations, and developer references.
- Fixed assertion failure messages for absent-entity checks to use the correct English plural:
  when count is 1 the message now says `"observed 1 matching workbook entity"` and when count is
  greater than 1 it says `"observed N matching workbook entities"`; the previous wording always
  used `"entities"` regardless of count.
- Fixed selector targeting error messages (`AssertionStep`, `MutationStep`, and `InspectionStep`
  validation) to report allowed and actual target types using stable protocol type IDs
  (`CELL_BY_ADDRESS`, `WORKBOOK_CURRENT`, etc.) instead of Java class display names
  (`CellSelector.ByAddress`, `WorkbookSelector.Current`); consumers no longer need to know the
  Java class hierarchy to interpret the rejection message.
- Fixed `--print-task-catalog --lookup <id>` and `--print-protocol-catalog --lookup <id>` single-
  entry responses to include `protocolVersion` as the first field, matching the envelope shape of
  the full-catalog and search responses; previously both lookup commands emitted a bare value
  object with no version identifier.
- Removed the `targetSelectorRule` reference from the protocol-catalog Discovery notes in
  `--help-protocol`; `targetSelectorRule` is no longer emitted by the catalog for any entry, so
  describing it in the help text was misleading.
- Added `totalCount` field to `CatalogSearchResult` (the `--print-protocol-catalog --search`
  response) so consumers can see the total number of returned matches without counting the
  `matches` array; the field is always equal to `matches.length` and appears before `matches` in
  the serialized JSON.
- Removed three redundant lines from the `Request` section of `--help-protocol`: the standalone
  `"do not send step.type."` sentence (already the closing clause of the step-kind explanation
  that immediately follows it), the `"The request JSON itself is capped at 16 MiB"` line
  (already present in the `Limits` section as the `Request JSON size` entry), and the
  `"source.type must be NEW; mutation actions limited to ENSURE_SHEET and APPEND_ROW"` line
  (already present in the `Limits` section as the `STREAMING_WRITE mode` entry).

### Security

- Named-range formula injection rejection (`LIM-032`): `NamedRangeTarget.Formula` now calls
  `FormulaInputSecurity.rejectDde`, blocking `DDE(` and `WEBSERVICE(` payloads; named-range
  formulas are stored in the xlsx defined-names part and auto-evaluated by Excel on open, making
  them as dangerous as cell-formula injection.

- Chart title formula injection rejection (`LIM-033`): `ChartTitleInput.Formula` now calls
  `FormulaInputSecurity.rejectDde`; chart title formulas are stored in the chart XML and evaluated
  automatically when Excel renders the chart, giving them the same auto-execute risk as cell formulas.

- UDF formula template injection rejection (`LIM-034`): `FormulaUdfFunctionInput.formulaTemplate`
  now calls `FormulaInputSecurity.rejectDde`; applied as defense-in-depth since POI's evaluator
  does not execute DDE or WEBSERVICE server-side and templates are not written to xlsx output.

- Documented three accepted limitations in `LIMITATIONS.md`: `LIM-035` (`HYPERLINK()` formula
  content not validated against the URL scheme allowlist — requires user click, not auto-executed),
  `LIM-036` (`RTD()` not blocked — Windows/COM only, no-op on macOS/Linux), and `LIM-037`
  (`WorkbookFactory` may open XLS files named `.xlsx` — relates to LIM-002; no injection risk).

- Source file path confinement (`LIM-030`): `SourceBackedPathResolver.resolvePath` now delegates
  to `ExecutionRequestPaths.normalizePath` (LIM-025, LIM-029), preventing relative path traversal
  via `TextSourceInput.Utf8File` and `BinarySourceInput.File` inputs; previously a caller could
  submit `"../../etc/shadow"` as a source file path and read arbitrary server-side files whose
  content would appear in cell values or formula strings returned in the response.

- WEBSERVICE formula rejection (`LIM-031`): `FormulaInputSecurity.rejectDde` now also rejects
  formulas beginning with `WEBSERVICE(` (case-insensitive, after stripping `=` or `{=…}` prefix);
  `WEBSERVICE` causes Excel to make outbound HTTP requests during formula evaluation, making it a
  data-exfiltration vector when the generated workbook is opened by downstream users.

- Relative path traversal prevention (`LIM-025`): `ExecutionRequestPaths.normalizePath` now
  verifies that relative paths resolve within the working directory; paths using `../` components
  to escape are rejected with `INVALID_REQUEST`. Absolute paths remain allowed as explicit
  references.

- DDE formula injection rejection (`LIM-023`): `CellInput.Formula` now rejects inline formula
  sources that begin with `DDE(` (case-insensitive, after stripping the leading `=`) to prevent
  malicious DDE calls from being written into saved workbooks.

- Explicit ZIP decompression limits (`LIM-026`): `ExcelWorkbookOpenSupport` now explicitly
  configures Apache POI `ZipSecureFile` with a 100 MiB maximum decompressed entry size and a
  1:100 minimum inflate ratio, replacing implicit reliance on POI defaults.

- Step count limit (`LIM-024`): `WorkbookPlan` now enforces a maximum of 10,000 steps per
  request, preventing unbounded resource consumption from pathological inputs that could otherwise
  exhaust execution time or memory under the 16 MiB JSON cap.

- DDE rejection extended to all formula inputs (`LIM-027`): `FormulaInputSecurity.rejectDde` now
  guards every formula-bearing input type — `ArrayFormulaInput`, all `DataValidationRuleInput`
  formula fields, `ConditionalFormattingRuleInput` formula fields,
  `ConditionalFormattingThresholdInput.formula`, and `ChartDataSourceInput.Reference.formula` —
  closing the gap left by the original LIM-023 fix which only covered `CellInput.Formula`.

- URL scheme allowlist for hyperlinks (`LIM-028`): `HyperlinkTarget.Url` now accepts only `http`,
  `https`, `ftp`, and `ftps` schemes; previously any non-`file`/`mailto` URI scheme was accepted,
  allowing `javascript:`, `vbscript:`, `ms-excel:`, and `ldap:` URLs to be written into workbooks.

- Symlink confinement detection (`LIM-029`): `ExecutionRequestPaths.normalizePath` now walks each
  path component and verifies that symbolic links within the working directory resolve to targets
  still inside it, closing the bypass of the lexicographic LIM-025 traversal check.

## [0.64.0] - 2026-05-08

### Changed

- Bumped `org.apache.santuario:xmlsec` from `3.0.6` to `4.0.4`, added the required Jakarta Activation and XML Bind APIs to the runtime surface, and replaced the OOXML signature relationship-transform path with a GridGrind-owned xmlsec-4-compatible DSig provider.
- Split the old contract-owned CLI help surface back into the `cli` module, promoted workbook/package-signing seams to explicit engine bridge types, and replaced the widened sheet-scoped `GET_CHARTS` inspection path with first-class chart selectors.
- Reduced the contract discovery layer to public task descriptors, moved task planning and English keyword query ranking into `cli`, and switched task-plan scaffolds to generic descriptor-derived starter requests instead of hidden contract-owned starter workflows.
- Rebuilt the protocol-catalog descriptor stack around lightweight canonical descriptor records, moved the request-execution bridge into `engine`, replaced the old `--print-goal-plan` surface with the honest `--print-task-keyword-match <query>` command and `TaskKeywordMatchReport`, and retargeted the build-failing architecture ratchet onto the current engine and Jazzer hotspot files.
- Split the old omnibus `InspectionResult` union into workbook-, sheet-, asset-, surface-, and analysis-scoped public result families, moved inspection-result null or blank guards into shared support, and registered JSON subtype ids from the sealed leaf result types instead of one central annotation block.
- Promoted formula-surface, sheet-schema, and named-range-surface reads into an explicit public `Surface` query family, renamed their response anchors from `analysis.*` to `surface.*`, moved the built-in example registry behind a non-exported CLI package, and split the internal example catalog into smaller discovery sub-contexts.
- Added `Woodstox Core` and `Stax2 API` to `NOTICE` and `PATENTS.md`; added `LICENSE-BSD-2-CLAUSE` to the distribution (Stax2 API is BSD 2-Clause); added `LICENSE-EDL-1.0` and `LICENSE-BSD-2-CLAUSE` to the Docker image filesystem; updated the `org.opencontainers.image.licenses` OCI label to `MIT AND Apache-2.0 AND BSD-2-Clause AND BSD-3-Clause AND EDL-1.0`; wired `LICENSE-BSD-2-CLAUSE` into the `--license` CLI output path.
- Replaced the hand-maintained dependency enumeration in `README.md`'s Legal section with a license-family summary and a pointer to `NOTICE` as the canonical attribution source; added `LICENSE-BSD-2-CLAUSE` and `LICENSE-EDL-1.0` links to the README legal footer.
- Made unbundled `:cli:run --version` resolve the packaged product version from the bundled resource metadata, taught the built-in example catalog and help output which requests are self-contained versus asset-backed, preserved actionable response-write failure messages instead of collapsing them to a bare path, and cleaned stale versioned jars out of module `build/libs` directories during fresh jar production.
- Replaced centralized protocol action/query/assertion targeting tables with leaf-owned metadata, removed the dead top-level descriptor registry classes left behind by the split, published exact `requiredPaths` for asset-backed built-in examples through the machine-readable example catalog and help surface, and corrected the example-guide verification command to the live `:cli:test` fixture owner.
- Reworked the root `check.sh` stage runner so stage completion follows the actual command process instead of a lingering `tee`/`awk` pipe, which removes the false Stage 4 shell-regression stalls that appeared on GitHub runners after the last regression script had already reported success.
- Extended the release merge-handoff verifier wait budget to outlast the full `main` `Gate` path, so release tagging no longer aborts while a healthy post-merge `Docker smoke` leg is still running on GitHub.
- Tightened the workbook-core null model by replacing internal null sentinels with `Optional` across pivot-table, drawing, validation, color, signature-line, and embedded-object seams; enabled NullAway in the shared Java conventions for `@NullMarked` production packages; and pinned JaCoCo to the exact published Maven snapshot coordinate (`0.8.15-20260506.113836-98`) that corresponds to the May 6 2026 trunk build required for Java 26 coverage.

### Fixed

- Non-success CLI runs that write their payload to `--response <path>` now emit one stderr line naming that response or doctor-report file, so operators and agents can find the structured failure payload without guessing after a non-zero exit.
- CLI parse and discovery failures now honor `--response <path>` the same way execution failures do, response-write fallback now emits one stderr line before streaming the structured failure to stdout, and the `Protocol Grammar` help block now keeps constructor shortcuts and playbook advice in the `Operator Guidance` section instead of mixing them into the normative request contract.
- Hardened the Stage 4 shell regression harnesses to the active shell baseline: Jazzer wrapper UX checks no longer depend on fragile direct redirection under the release-contract prelude, the wrapper arg-order regression no longer assumes Bash `mapfile`, the root `check.sh` monitor path no longer depends on process substitution that could terminate Docker-smoke verification early, and the deterministic root gate now forces serial Gradle execution after a fresh parallel release build produced a truncated JPMS jar entry.

## [0.63.0] - 2026-05-05

### Changed

- Bumped `tools.jackson.core:jackson-databind` from `3.1.2` to `3.1.3`; annotations remain on the Jackson 2.x coordinates at `2.21` by design.
- Hardcoded the release-blocking check in `verify-release-merge-handoff.sh` to `Gate` (the single aggregate CI check) and removed the `GRIDGRIND_RELEASE_BLOCKING_CHECKS` env-var override; the previous default included `Check,Docker smoke,Contributor devcontainer`, which is inconsistent with the already-hardened `verify-release-candidate-tag.sh`.
- Updated `test-verify-release-merge-handoff.sh` and `test-verify-release-candidate-tag.sh` regression tests to provide `Gate` as the single successful check run and `Gate` failure as the failure case, replacing the old three-check format.
- Suppressed `[GRADLE-TEST-PULSE]` and `[JAZZER-PULSE]` lines from `check.sh` stdout; they are still written to the stage log for post-run inspection but no longer flood the terminal during a run.

### Fixed

- Pinned all GitHub Actions workflow runners from the floating `ubuntu-latest` label to `ubuntu-24.04` across all four workflows (`ci.yml`: 3 jobs; `release.yml`: 1 job; `container.yml`: 2 jobs; `gradle-wrapper-validation.yml`: 1 job) so runner image updates cannot silently change the build or release environment.
- Added `workflow_dispatch:` trigger to `ci.yml` so maintainers can manually rerun the aggregate `Gate` against a branch when GitHub fails to attach the `pull_request` workflow on initial PR open.
- Added a `devcontainer-changes` path-detection job to `ci.yml` that computes a git diff against the devcontainer trigger paths (`.devcontainer/`, `scripts/validate-devcontainer.sh`, `scripts/devcontainer-prepare-user-home.sh`); non-devcontainer PRs skip the full Docker build-and-validate cycle, reducing typical PR wall-clock time significantly.
- Added a `Gate` aggregate required-status job to `ci.yml` using `if: always()` with explicit `${{ toJSON(needs.*.result) }}` failure detection so a correctly skipped `devcontainer` gate does not prevent `Gate` from being reported or block merge — only a failed or cancelled job prevents success. Configure branch protection to require `Gate` as the single required check.
- Promoted `permissions: contents: read` to the workflow level in `ci.yml` and removed the redundant per-job declarations.
- Fixed `release.yml` concurrency to `cancel-in-progress: false`; the previous `true` setting could cancel an in-progress release publication mid-way if a second tag was pushed, leaving the release in a half-published state.
- Raised `container` job `timeout-minutes` from 20 to 25 to provide a clear margin between Docker image build-and-push time and the post-publish verification step.
- Added `id-token: write` permission to the `container` job in `container.yml` to activate the OIDC token flow required for keyless provenance attestation signing; the job already enabled `provenance: mode=max` and `sbom: true` but lacked the permission to complete attestation signing.
- Updated `actions/checkout` in `gradle-wrapper-validation.yml` from `v5.0.0` (`08c6903cd8c0fde910a37f88322edcfb5dd907a8`) to `v6.0.2` (`de0fac2e4500dabe0009e67214ff5f5447ce83dd`) for consistency with all other workflows; added `timeout-minutes: 10`.
- Hardcoded the release-blocking check in `verify-release-candidate-tag.sh` to `Gate` (the single aggregate CI check) and removed the `GRIDGRIND_RELEASE_BLOCKING_CHECKS` env-var override; the previous default included `Contributor devcontainer` which is legitimately skipped on commits that do not touch devcontainer files, causing false failures on those release commits.
- Added retry support to `verify-github-release.sh` (default 3 attempts with 5-second delay, overridable via `GRIDGRIND_GITHUB_RELEASE_VERIFY_RETRIES` and `GRIDGRIND_GITHUB_RELEASE_VERIFY_DELAY_SECONDS`) so release asset availability checks are resilient to brief GitHub API propagation lag.
- Added `isReproducibleFileOrder = true` to the `shadowJar` task in `cli/build.gradle.kts` for deterministic, auditable JAR entry ordering across builds.
- Added `jazzer` Gradle ecosystem and `docker` ecosystem entries to `.github/dependabot.yml`; the nested `jazzer/` Gradle build was not scanned for dependency updates, and the Dockerfile base image had no automated update coverage.
- Updated `RELEASE_PROTOCOL.md` to replace all references to the old three-job blocking check list (`Check`, `Docker smoke`, `Contributor devcontainer`) with the new single `Gate` aggregate check.
- Added a Dependabot Approval Strategy section to `RELEASE_PROTOCOL.md` with triage tiers (security within 7 days, regular before next release, major version bumps as considered upgrades), required CI gates before any merge, and explicit prohibitions.
- Fixed `nested/ plain` typo in the `--operation <id>` flag description caused by a Java string-concatenation line split at the slash; the rendered text now correctly shows `nested/plain`.
- Added a separator line (`------- --------------------`) between the column headers and rows of the `Coordinate Systems` table so the header is visually distinct from the coordinate data.
- Removed a dangling sentence fragment ("Explicit formula-calculation policy.") from the `execution.calculation` bullet in the `Request` section; the bullet now reads as a single coherent description that flows from the field name through the policy variants without an orphaned nominal clause.
- Fixed bare invocation (`gridgrind` with no arguments) producing a missing trailing newline on the help output, causing zsh and other POSIX-compliant shells to display a `%` prompt indicator; the implicit-help path now routes through the same `CliResponseWriter.writePayload` infrastructure as `--help` so the newline guarantee is uniform across all help entry points.
- Moved the `Flags:` section immediately after `Usage:` so the CLI grammar surface (usage patterns and flag definitions) is presented as a contiguous block before operator-guidance sections such as `Workflows:`, `Execution:`, `Limits:`, and `Request:`; this follows the conventional `--help` structure and prevents operator-guidance prose from appearing to be first-class grammar.
- Renamed CLI help section `First-Contact Workflows` to `Workflows` to use cleaner Ubiquitous Language that does not expose internal jargon to operators and agents.
- Renamed CLI help section `stdin Example` to `Stdin Example` for consistent Title Case across all section labels.
- Renamed CLI help section `Docker File Example` to `Docker Example`; the previous name implied a Dockerfile definition rather than a `docker run` command example.
- Moved the `ANALYZE_WORKBOOK_FINDINGS` prose summary out of the Discovery command list into the prose note at the bottom of the Discovery section; it was rendered at the same indent and visual level as `gridgrind` commands, making it look like an invocable CLI entry.
- Indented the built-in generated examples list one additional level (from 2 to 4 spaces) so the entries appear visually nested under the `Built-in generated examples:` sub-label rather than at the same indent as the CLI commands above them.
- Replaced the implicit `Column structural edits: same ownership rule` cross-reference with the explicit rule text so the Limits entry is self-contained and readable without requiring the adjacent `Row structural edits` entry for context.

## [0.62.0] - 2026-05-01

### Changed

- The shared version catalog now pins JUnit `6.1.0-RC1` and the exact published JaCoCo snapshot
  artifact that corresponds to official trunk build `0.8.15.202604290352`, keeping Java 26
  coverage support aligned with a concrete reproducible Maven coordinate.
- Problem diagnostics and response reports now live in focused public contract namespaces instead
  of two giant wire god-files, so execution-stage contexts, persistence outcomes, workbook facts,
  layout facts, schema or formula facts, analysis facts, and failure facts each have one narrower
  owner.
- Workbook coordination now flows through one immutable workbook context instead of direct
  controller hoarding inside `ExcelWorkbook`, which reduces cross-feature coupling at the
  workbook façade boundary.
- The chart authoring contract moved JSON-entry defaults behind owned normalization helpers so
  public `@JsonCreator` methods no longer carry ad hoc defaulting branches inline.

### Fixed

- The packaged `installShadowDist` launcher now ships as `gridgrind`, and its first-contact help
  path no longer leaks Java native-access warnings or other stderr noise before the product help
  text.
- CLI discovery failures now explain their next corrective move directly: dependent flags such as
  `--task`, `--operation`, and `--search` name the parent command they require, unknown task ids
  point back to `--print-task-catalog` and `--print-task-keyword-match`, unknown protocol lookup ids point
  to protocol-catalog search, and missing request files identify the unreadable path explicitly.
- Built-in example lookup failures now point operators back to the packaged help surface’s
  generated-example section, and the packaged help synopsis now advertises `--license` alongside
  the other first-contact commands.
- Inserted rows inside existing sheet content now inherit adjacent visual formatting instead of
  materializing as workbook-default blanks, so row-height, row-style, font, fill, wrap, and cell
  style facts follow the surrounding authored sheet after `INSERT_ROWS`.
- Inserted columns inside existing sheet content now inherit adjacent visual formatting instead of
  materializing as workbook-default blanks, so column-width, default-column-style, font, fill,
  wrap, and cell-style facts follow the surrounding authored sheet after `INSERT_COLUMNS`.
- Cell text, workbook style, worksheet hyperlink, formula length, formula nesting, and formula
  argument ceilings from the limitations registry are now anchored at central write seams with
  explicit enforcement instead of being left implicit in downstream Apache POI failures.
- Chart-title live-resolution fallbacks no longer flood fuzzing sessions with repeated warnings
  when a title formula evaluates to an error-cached scalar and the code intentionally falls back
  to cached or empty text.
- Data-validation request payloads now require explicit `allowBlank` and
  `suppressDropDownArrow` booleans on the public wire surface instead of inheriting hidden JSON
  defaults, and the shipped request examples and Jazzer regression fixtures now author those
  choices directly.
- Jazzer verification tool tasks now keep their arguments lazy without forfeiting configuration
  cache safety, so `jazzerStatus`, `jazzerReplay`, `jazzerRefreshPromotedMetadata`, and Gradle
  task listing no longer discard the cache or demand unrelated Gradle properties.
- Jazzer wrapper `--help` and invalid-target failures now stay on the project-owned wrapper
  surface instead of falling through to generic Gradle help or Java stack traces.
- Jazzer replay and promotion wrappers now reject missing or non-file inputs before Gradle boots,
  and invalid replay or promotion targets now print the correct wrapper usage instead of leaking
  the internal `_run-task` contract.
- Jazzer status now keeps a stable `Jazzer Status` surface even on a fresh workspace with no
  recorded summaries, the documentation-contract regression no longer depends on `rg` being
  installed on the host shell, and the release-surface container verifier now drives
  `--doctor-request` from the printed request template instead of a drift-prone hard-coded stub.
- `STREAMING_WRITE` now keeps SXSSF shared strings enabled, so repeated-text workbooks no longer
  balloon into inline-string-heavy `.xlsx` packages that are much larger than equivalent
  full-XSSF or Excel-authored files.
- Built-in and checked-in example requests now save into request-local `generated-workbooks/`
  paths instead of hard-coding the repository build tree into printed example payloads.
- Request JSON once again accepts the documented omitted top-level defaults for
  `protocolVersion`, `persistence`, `execution`, `formulaEnvironment`, and `steps`, while
  continuing to reject explicit `null` placeholders on the public wire.
- Comment and signature-line request payloads now require explicit authored booleans on the wire
  instead of silently coercing missing visibility or comment-allowance fields to defaults during
  JSON decoding.

## [0.61.0] - 2026-04-29

### Changed

- The Gradle wrapper now pins stable Gradle `9.5.0`, and the shared build conventions now add the
  official JaCoCo snapshots repository narrowly for `org.jacoco` artifacts so the repo can pin
  the exact published snapshot that corresponds to JaCoCo trunk build
  `0.8.15.202604281210` while keeping the rest of dependency resolution release-only.
- Request, response, and discovery JSON readers now reject explicit `null` placeholders anywhere
  in the public wire surface. Absent state must be expressed by omitting the property, so
  GridGrind no longer accepts `null` as an alternate control channel during request parsing or
  machine-readable catalog/report reads.
- Selector is now a sealed root with explicit permitted selector families, so selector dispatch
  and selector-oriented tests regain compile-time exhaustiveness instead of relying on ad hoc
  fallback implementations.
- The contributor docs now include a step-by-step Docker-only Jazzer walkthrough for first-time
  Docker users, including image build, container entry, the first short active harness run, the
  output to expect while fuzzing, and where to inspect the resulting summaries afterward.
- The contributor docs now treat the Dev Container Specification as the canonical environment
  owner, keep VS Code as an optional overlay rather than a hard dependency, and document the
  tooling-agnostic `devcontainer up` / `devcontainer exec` workflow step by step alongside the
  existing integrated and raw-Docker paths.

### Fixed

- Conditional-formatting differential colors, borders, and data-validation prompt/error report
  fields, plus dynamic autofilter numeric bounds, now use typed absence in the contract layer
  instead of raw `null` padding, so Java authoring, executor converters, and fuzz invariants
  share the same presence model as the JSON wire format.
- Missing-cell, missing-workbook, and unregistered-UDF failures no longer masquerade as
  `IllegalArgumentException`, unchecked IO wrapping now preserves IO failure type, and chart-title
  formula resolution failures now leave warning-level evidence instead of disappearing silently.
- The nested Jazzer support-test and coverage gates now bind deterministic `*Test.class` inputs
  explicitly under Gradle `9.5.0`, so the Jazzer verification stage no longer compiles its
  test suite and then misclassifies it as `NO-SOURCE`.
- The supported `jazzer/bin/*` wrapper surface now forwards Gradle properties and console options
  in the order Gradle actually accepts, so documented commands such as
  `jazzer/bin/fuzz-protocol-request -PjazzerMaxDuration=5m --console=plain` start live fuzzing
  instead of falling back to Gradle help.
- Terminal-only Docker fuzzing now uses cross-platform corpus-size accounting in the Jazzer
  launcher, so active harness commands no longer fail immediately inside the Linux contributor
  container while the same command works on macOS.
- The committed devcontainer now repairs its named Gradle and general-cache mounts on start, and
  devcontainer validation now proves those mounts stay writable for the remote user even after a
  prior ad hoc Docker workflow left root-owned entries behind.
- Source-backed request resolution no longer leaves selector and structured-authored-value binding
  collapsed into one PMD-suppressed seam. Selector resolution and structured payload resolution
  now live in dedicated helpers, keeping the top-level resolver closer to orchestration than a
  kitchen-sink translator.
- Structured workbook-command translation no longer keeps sheet-layout, drawing, and tabular
  authored-input conversion collapsed into one oversized PMD-suppressed helper. Dedicated layout,
  drawing, and tabular converters now carry those seams, keeping the remaining validation/filter
  converter small enough for the enforced import and size audits.
- Read-request failures now preserve a path-only JSON cursor when semantic validation can still
  identify the offending request field but no stable line/column survives tree-backed parsing, so
  CLI and response contexts no longer erase that cue entirely.
- The repo now carries a permanent `.codex/OBSERVATIONS_INCIDENTAL.txt` audit ledger, Gradle
  dependencies are covered by Dependabot, and stale investigation or release artifacts under
  `tmp/` are cleaned instead of lingering between sessions.

## [0.60.0] - 2026-04-28

### Changed

- GridGrind now ships a committed contributor devcontainer based on a pinned glibc image plus
  Azul Zulu 26, and the developer docs now treat that containerized workflow as the preferred
  contributor path while keeping host-native Java guidance accurate as a fallback.
- The shared Gradle build logic now uses Kotlin `2.4.0-Beta2` on Gradle `9.5.0-rc-4` and emits
  JVM 26 bytecode directly, so the build no longer carries a separate JVM 25 exception below the
  repository's Java 26 baseline.
- The canonical `./check.sh` gate now runs with an isolated repo-scoped `GRADLE_USER_HOME`,
  always forces `--no-daemon`, and shares one repo-wide verification lock with devcontainer,
  Docker-smoke, and Jazzer wrapper entrypoints so local full verification no longer overlaps
  through shared daemon or cache state by accident.
- `./check.sh` now derives its fixed five-stage inventory, Stage 4 release-surface shell
  regression list, and usage text from `scripts/check-stage-contract.sh`, and a dedicated shell
  regression now guards that canonical owner so the root gate cannot drift back into parallel
  stage definitions.
- Discovery and planning JSON now omit absent optional fields instead of emitting explicit
  `null` placeholders. `--print-protocol-catalog`, `--print-task-catalog`,
  `--print-task-plan`, `--print-task-keyword-match`, and `--doctor-request` stay schema-compatible while
  becoming easier for agents and shell tooling to consume directly.
- Request and response JSON now omit absent optional fields too, including `--print-example`
  output, checked-in `examples/*.json`, and normal execution responses, so the full public JSON
  surface no longer uses explicit `null` placeholders as a control protocol.
- Workbook-analysis finding codes and severities now come from one shared
  `excel-foundation` owner instead of parallel copies in engine and contract modules, so the
  published vocabulary no longer depends on lockstep duplicate enums.
- The limitations registry now distinguishes product-enforced limits from upstream reference
  ceilings instead of implying that every documented Excel or POI ceiling is surfaced through the
  same runtime/help/catalog propagation path.
- Canonical request plans now keep default execution and formula-environment state as normalized
  objects in memory instead of collapsing those sections back to `null`, so authoring, executor,
  and JSON emission all share one default model.
- Request-doctor and calculation telemetry contracts now keep optional summary, problem,
  preflight, and message state as explicit `Optional`-backed fields in memory instead of raw
  `null` padding, so Java authoring, CLI fallback handling, executor telemetry, and tests all
  share the same absence model as the public JSON wire shape.
- Jazzer's deterministic coverage gate now treats the extracted `OperationSequence*` generator
  helpers as one exclusion family, so splitting the fuzz-workflow generator into smaller seams no
  longer breaks Stage 2 verification just because the coverage scope list drifted.

### Fixed

- CI now validates the committed devcontainer surface in addition to the whole-repo deterministic
  gate, so the preferred contributor environment cannot silently drift away from the documented
  Java, editor-routing, and container-base contract.
- Release CI no longer strands branch protection on stale check names while the release verifiers
  silently ignore the committed devcontainer job. The `CI` workflow now keeps `Check` and
  `Docker smoke` as the required contexts, exposes `Contributor devcontainer` as a separate job,
  and the tag/merge release verifiers now treat that contributor-environment job as
  release-blocking too.
- `./check.sh` progress monitoring no longer assumes BSD `stat`. The root gate now sources a
  portable file-size helper with BSD, GNU, and generic fallback behavior, and a dedicated shell
  regression keeps the Linux/macOS seam from breaking release CI again.
- The CLI-contract and public-container publication regressions no longer push full fixture
  payloads through huge process environments. They now write case fixtures under `tmp/` and
  verify those release surfaces through file paths, avoiding Linux `Argument list too long`
  failures in release CI.
- Release-surface shell regressions now follow the repo search-tool portability contract: they
  prefer `rg` when available but fall back cleanly when it is not, instead of silently weakening
  or breaking verification on lean CI runners.
- The engine now names every private Apache POI contract it still depends on for relation
  removal, sheet-clone hyperlink preparation, workbook picture-catalog synchronization, and
  gradient-fill registry access. Compatibility tests and runtime failures now point at the exact
  broken POI seam and affected feature instead of failing later with generic reflection errors.
- `--print-protocol-catalog` no longer routes full-catalog, ranked-search, and exact-entry lookup
  through one nullable union command model. The CLI now uses dedicated typed variants internally,
  so discovery behavior is explicit instead of being inferred from null combinations.
- Protocol-catalog target-selector metadata no longer relies on internal null-sentinel helper
  returns. Derived selector rules now flow through typed `Optional` seams, and exact entry lookups
  stop printing `targetSelectorRule: null` noise for operations that do not need a rule.
- Release-surface shell regressions no longer validate the CLI contract against hand-maintained
  fake help/catalog/task/goal/doctor blobs. The verifier tests now derive their baseline payloads
  from the built GridGrind CLI artifact and mutate only the field under test.
- Current agent and Jazzer architecture guidance no longer teaches the deleted top-level
  `protocol` module, and a build-failing architecture audit now guards the live
  `cli -> executor -> contract` plus `executor -> engine` graph.
- The success/failure response seam no longer depends on nested null-padding helpers for cell and
  problem-context variants. `CellReport` and `ProblemContext` now live as dedicated top-level
  contract types, and a build-failing seam audit guards `GridGrindResponse.java` against growing
  back into an unbounded god-file.
- Build-failing seam audits now also guard `ProblemContext.java`,
  `OperationSequenceModel.java`, and `OperationSequenceValueFactory.java` after splitting their
  selector/assertion and chart-generation sub-foundations into dedicated helpers, so contract and
  Jazzer boundary theory no longer depends on a handful of oversized god-files.
- Chart clone-preparation helpers no longer use raw `null` to represent missing formula-node text
  payloads. The rewrite seam now uses `Optional` internally and preserves the same public failure
  when POI loses the node text during restore.

## [0.59.0] - 2026-04-25

### Changed

- Structured success and failure response construction now uses explicit protocol factories for
  synthetic journal-backed results, instead of public convenience constructors that inferred
  omitted execution details from `null`.
- Presence assertions are now family-specific on the public wire surface. `EXPECT_PRESENT` and
  `EXPECT_ABSENT` have been replaced by `EXPECT_NAMED_RANGE_PRESENT`,
  `EXPECT_NAMED_RANGE_ABSENT`, `EXPECT_TABLE_PRESENT`, `EXPECT_TABLE_ABSENT`,
  `EXPECT_PIVOT_TABLE_PRESENT`, `EXPECT_PIVOT_TABLE_ABSENT`, `EXPECT_CHART_PRESENT`, and
  `EXPECT_CHART_ABSENT`, so selector intent is explicit in the authored request instead of being
  recovered from ambiguous shared selector ids.
- Generated CLI help and the machine-readable CLI catalog now publish first-contact workflows for
  discovery, drafting, doctoring, and execution, instead of presenting the public surface only as
  a flat flag inventory.
- Generated CLI help now teaches the save-then-reopen workflow explicitly too, including
  `source.type=EXISTING` for reopening an `.xlsx` workbook from disk and the matching
  `source.path` file-workflow wording, so first-contact operators do not have to infer the
  existing-workbook discriminator from deeper protocol material.
- Generated CLI help, the quick reference, and the request/execution reference now call out that
  every non-empty step needs a caller-defined `stepId`, so first-contact request authors do not
  have to discover that requirement only from a shape error after writing their first mutation or
  inspection.

### Fixed

- Engine and executor helper seams no longer rely on test-only deep reflection or null-sentinel
  returns for key optional facts. Residual coverage now exercises package-visible seams directly,
  autofilter sort-state readback uses typed absence internally, and executor warning, diagnostic,
  and analysis-severity helpers no longer encode "missing" with raw `null` in business logic.
- Literal cell writes now actually replace existing formula cells instead of only updating POI's
  cached formula result. `SET_CELL` and `SET_RANGE` clear the old formula first, so writing a
  number, text, boolean, blank, date, or date-time value no longer leaves the prior formula alive
  under the new cached display value.
- `INSERT_ROWS` and `INSERT_COLUMNS` no longer fail just because the target sheet already owns
  data validations. GridGrind now snapshots those validation structures before the structural
  insert, shifts their covered ranges and supported formulas authoritatively, and reapplies them
  after the insert so validated workbooks can be extended without a manual clear-and-rebuild pass.
- `COPY_SHEET` no longer crashes when Apache POI tries to clone charts whose data sources are
  authored through workbook or sheet defined names. GridGrind now normalizes those chart formulas
  to explicit area references just for the clone pass, restores the source-sheet chart formulas
  immediately afterward, and leaves the copied chart bound to the copied sheet instead of failing
  with a negative-column reference error.
- Failure diagnostics now require an explicit `problem.causes[*].stage` value, so structured cause
  entries always preserve the pipeline stage that classified the failure.
- Contract-exposed foundation enum tokens and gradient-fill type tokens are now pinned by
  regression coverage, so accidental wire-vocabulary renames fail verification instead of
  silently drifting the published protocol surface.
- `./check.sh` Stage 4 documentation and usage text now list `scripts/test-verify-cli-contract.sh`
  alongside the other release-surface shell regressions that the script already executes.
- `--doctor-request` now accepts `--response <path>` just like normal execution, so machine-
  readable doctor reports can be written to files in both fat-JAR and Docker runs instead of
  forcing stdout-only handling.
- `--response <path>` now works consistently across discovery and informational CLI commands too,
  including `--help`, `--version`, `--print-request-template`, `--print-example`,
  `--print-task-catalog`, `--print-task-plan`, `--print-task-keyword-match`, and
  `--print-protocol-catalog`, with the same parent-directory creation and structured fallback
  behavior already used by execution and doctoring.
- Missing polymorphic `type` discriminators in request JSON no longer crash the parser's public
  error-reporting path. Requests that omit a step assertion, action, query, source, or other
  subtype `type` field now fail deterministically with the product-owned
  `Missing required field 'type'` contract instead of surfacing a null-discriminator
  `NullPointerException`.
- Public Docker-first guidance now calls out that a plain `docker run ...:latest` can reuse a
  stale local `latest` tag. Documentation now tells operators to `docker pull` first or use
  `--pull=always` when they need the registry's current `latest`.
- User-facing Docker-first copy-paste examples now use the freshness-safe form too. Documentation
  no longer warns about stale `latest` tags in prose while still demonstrating plain
  `docker run ...:latest` as the first command a new operator is likely to copy.
- First-contact request-shape failures now teach the intended replacement surface instead of only
  echoing an unknown discriminator token. `source.type=FILE` now points operators at
  `source.type=EXISTING`, the removed ambiguous assertion ids point at the new family-specific
  assertion ids, and `--print-example` now suggests the stable upper-case example id when the
  authored token matches a shipped example file stem such as `chart-request`.

## [0.58.0] - 2026-04-24

### Changed

- Public docs now distinguish built-in examples that are fully self-contained in a blank artifact
  workspace from the three repo-asset-backed examples (`CUSTOM_XML`, `SOURCE_BACKED_INPUT`, and
  `PACKAGE_SECURITY_INSPECTION`) that require copied `examples/` assets. Runtime example tests now
  cover that portability split directly instead of leaving it implied by prose.
- Public docs and generated help now describe the doctor surface more precisely:
  `--doctor-request` validates request shape, source-backed authored input resolution, and
  existing-workbook source accessibility up front, while still leaving workbook mutation to the
  real execution path.
- Built-in example discovery now publishes machine-readable workspace requirements through
  `workspaceMode` and `requiredPaths`, so operators and agents can tell which examples need copied
  repository assets without relying on prose or trial execution.
- The release protocol now explicitly treats mounted or removable primary checkouts that break
  Gradle locking or stall preflight as problematic release hosts and directs operators onto a clean
  local worktree before any release build, instead of leaving that failure mode implicit.

### Fixed

- `SAVE_AS` persistence now creates missing parent directories consistently on both the normal
  XSSF path and the `STREAMING_WRITE` path, matching the documented file-workflow contract.
- `--doctor-request` no longer reports some requests as valid when they would immediately fail
  during input resolution or workbook opening at execution time. Source-backed authored input
  loading and existing-workbook accessibility are now preflighted in doctor mode as well.
- Bare `help` now works as a first-class alias for `--help` in the packaged CLI instead of failing
  as an unknown argument.
- Non-step CLI failures now preserve their real problem classification in the synthetic execution
  journal instead of always backfilling `INTERNAL_ERROR`.
- Root-project formatting checks no longer race concurrent compilation by scanning a mutating
  checkout during `./check.sh`; the repository-wide Spotless project-file pass now runs against a
  stable tree before build outputs are produced.

## [0.57.0] - 2026-04-24

### Fixed

- Running the packaged CLI with no arguments from an interactive terminal no longer hangs silently
  while waiting for stdin. `gridgrind` and `java -jar gridgrind.jar` now print the same help text
  as `--help` and exit with code `0`, including cases where stdin is interactive but stdout is
  redirected. The release-surface CLI verifier now asserts that interactive no-arg behavior under
  a real PTY so this path cannot regress quietly again.
- The Java authoring surface now stops cleanly at the canonical `WorkbookPlan` boundary instead of
  pulling executor runtime concerns into the `authoring-java` module. `dev.erst.gridgrind.authoring`
  now depends transitively on `contract`, the fluent API no longer exposes `run(...)`, raw
  request DTOs, or response-report types as intended public authoring vocabulary, and the
  shipped Java example now executes in process through an explicit executor call.
- Task keyword match reporting is now less noisy for agents and operators. `--print-task-keyword-match` no longer keeps
  low-signal candidates alive solely because one capability summary happened to share a word with
  the goal, and `suggestedIntentTags` now come from the strongest matching tasks instead of from
  the entire catalog.
- Protocol-catalog field-group metadata is no longer collapsed into one giant nested/plain support
  file. The nested and plain descriptor registries now live in dedicated support types again, with
  the thin facade left in place only as the local assembly seam.
- Release-surface shell regressions no longer duplicate the same fake help/catalog/task/doctor
  payloads in multiple scripts. The shared fake public-contract fixtures now live in one sourced
  shell helper so verifier updates have one maintenance point instead of two drifting copies.

## [0.56.0] - 2026-04-23

### Changed

- Task discovery now ships broader runnable starter scaffolds instead of thin placeholders.
  `--print-task-catalog`, `--print-task-plan`, and `--print-task-keyword-match` now cover tabular
  reports, dashboards, data-entry workflows, workbook audits, custom XML workflows, pivot
  reports, drawing/signature workflows, and workbook-maintenance flows with non-empty starter
  plans.

### Fixed

- CLI request-path handling is now portable and internally consistent. When the CLI reads a
  request via `--request`, relative request-owned paths now resolve from the request file
  directory instead of the shell working directory, including `source.path`, persistence paths,
  source-backed `UTF8_FILE` / `FILE` payloads, formula external workbook bindings, and OOXML
  signing keystore paths.
- CLI discovery commands now reject stray trailing flags consistently instead of silently
  accepting malformed invocations. `--help`, `--version`, `--license`, `--print-example`,
  `--print-task-catalog`, `--print-task-plan`, `--print-task-keyword-match`, `--print-request-template`,
  and `--print-protocol-catalog` now fail fast on extra arguments, and `--doctor-request` no
  longer ignores unknown trailing flags.
- Public docs, help text, and shipped examples now match the live discovery and path-resolution
  surface. The request/execution reference, quick-start flow, quick reference, README, operations
  index, and generated example set now document request-file-rooted relative paths, the broader
  task planner surface, and the new sheet-maintenance example.
- Release-surface shell validation no longer assumes every verification checkout is a live Git
  worktree. The Jazzer public-surface regression still verifies Git tracking when `.git` is
  present, but copied validation checkouts now pass based on the actual shipped shell surface
  instead of failing on missing repository metadata, and `./check.sh` now truthfully lists that
  Jazzer public-surface regression in its own Stage 4 contract summary.
- The largest new discovery seam no longer lives in one 900-line registry file. Task-definition
  builders are now split by responsibility, and the executor conversion surface remains split
  across workbook, cell, drawing, structured-feature, read-result, and workbook-report seams
  instead of regressing into new god-files.

## [0.55.0] - 2026-04-23

### Fixed

- Copied-sheet picture readback no longer depends on Apache POI leaving cloned drawing relation ids
  intact. GridGrind now repairs copied picture `r:embed` references against the target drawing
  relations after `COPY_SHEET`, picture reads or payload extraction now fail with a clear
  picture-specific integrity error instead of a raw `NullPointerException` when an external
  workbook is malformed, and the reproduced Jazzer `COPY_SHEET` picture crash is now covered by a
  promoted round-trip success seed plus persisted drawing-picture OOXML invariants.
- Malformed picture XML that still has `<pic:blipFill>` but no nested `<a:blip>` no longer leaks
  Apache POI's raw `NullPointerException` through picture inspection or delete flows. GridGrind
  now short-circuits those states as missing image relationships, keeps the delete path safe, and
  covers the blip-less picture seam with focused engine regressions.
- Copied embedded objects now keep their preview-image drawing relations intact through
  `COPY_SHEET` round-trips. GridGrind now repairs the preview media relationship referenced from
  the copied object shape itself, so saved copied worksheets no longer reopen with unresolved
  drawing `a:blip` ids in the embedded-object preview path.
- Copied embedded objects now also keep their worksheet-bound relation ids authoritative through
  `COPY_SHEET`, even when cloned `oleObject` or `objectPr` ids collide with the copied sheet's
  drawing, comment-VML, or other worksheet relations. GridGrind now rewrites those copied ids onto
  safe worksheet-owned relation ids, restores the sheet-to-drawing relation explicitly, treats
  sheet preview refs as image-only during read/delete flows, extends persisted OOXML invariants to
  worksheet `objectPr` preview refs, and promotes the reproduced round-trip artifact as a
  committed success seed.

## [0.54.0] - 2026-04-23

### Fixed

- The public contract no longer depends on the POI-backed engine module for shared Excel domain
  types. GridGrind now ships a dedicated `excel-foundation` module and package for the shared
  `.xlsx` enums, limits, and span/value objects consumed by the contract, engine, executor, and
  tests, so the `contract` module is a real boundary instead of a named-module cycle hiding behind
  `dev.erst.gridgrind.excel`.
- Developer architecture docs and release-surface shell regressions now describe and enforce the
  real six-module product graph. `excel-foundation` is now part of the documented JPMS graph, the
  accepted contract-replacement ADR no longer claims GridGrind is still a five-module system, and
  `./check.sh` no longer fails Stage 4 just because the shared foundation module exists.
- Formula authoring and formula-rewrite paths now go through one engine seam instead of each
  mutation mode setting formulas ad hoc. Streaming append writes, direct cell writes, copied-sheet
  sheet-name rewrites, and copied-table structured-reference rewrites now share consistent error
  handling for authored, rewritten, and scratch validation formulas.
- Protocol discovery no longer forces operators and agents to dump the full catalog and grep it
  externally. `gridgrind --print-protocol-catalog --search <text>` now returns ranked
  machine-readable matches across ids, qualified ids, groups, and summaries, and
  `--print-protocol-catalog` now rejects stray trailing flags instead of silently ignoring them.
- Copy-sheet table normalization is now deterministic across both in-memory mutation and later
  sheet rename flows. GridGrind now normalizes POI-cloned calculated-column body formulas back to
  its metadata-owned table shape, rewrites copied totals-row structured references from POI's
  transient clone names such as `Table2` to the final copied table names, and keeps copied sheets
  rename-safe instead of failing with structured-reference parser crashes or dead transient table
  names.
- Drawing media and container signature-line authoring are now reliable across the packaged fat
  JAR and Docker image. GridGrind now rebuilds POI's private workbook picture catalog from actual
  `/xl/media/*` parts before picture-backed mutations so signature-line preview images cannot
  collide with later picture or embedded-object preview creation, and the Docker image now ships
  fontconfig plus DejaVu fonts so `SET_SIGNATURE_LINE` works in headless container runs. Docker
  smoke now verifies signature-line authoring directly.
- The largest remaining protocol and Jazzer god-files are now split at stable seams instead of
  continuing to accumulate unrelated responsibilities. Nested/plain protocol field-group catalog
  descriptors now live in dedicated support types, Jazzer workbook-operation generation now
  separates orchestration from value-payload factories, and build-failing architecture audits now
  keep contract/foundation boundaries, formula-write centralization, POI private-access
  centralization, and the new file-size ceilings from silently regressing.
- Jazzer verification now keeps its coverage contract aligned with the intended deterministic
  support subset after the operation-generator split. The extracted
  `OperationSequenceValueFactory` now shares `OperationSequenceModel`'s coverage scope, and new
  selector-sweep seam tests directly exercise the extracted value factory plus workflow/command
  cleanup paths so the refactor stays verified without distorting the Jazzer coverage gate.

## [0.53.0] - 2026-04-22

### Fixed

- Public Markdown docs are now organized as stable entry-point docs plus focused long-form
  references. `docs/OPERATIONS.md` and `docs/QUICK_REFERENCE.md` are now kept intentionally small,
  the quick-start flow now teaches artifact-native `--print-example BUDGET`, the error reference
  now documents the live `CELL_NOT_FOUND` path truthfully, the Java 26 workstation guide now
  reflects the current official OpenJDK 26.0.1 download surface, and docs tests now fail on broken
  local Markdown links or if the entry-point reference docs bloat back into god-files.
- Worksheet comments now survive structural column edits deterministically even across the Apache
  POI 5.5.1 collision case where moved comments can reopen with duplicate persisted comment refs
  and the wrong visible note after `INSERT_COLUMNS`, `DELETE_COLUMNS`, or `SHIFT_COLUMNS`.
  GridGrind now rewrites the affected comment/VML state authoritatively after those edits, the
  discovered `.xlsx` round-trip finding is promoted as a committed success seed, and the Jazzer
  round-trip verifier now fails saved workbooks whose persisted comment refs are not unique.

## [0.52.0] - 2026-04-22

### Fixed

- Chart readback now resolves `XSSFGraphicFrame -> chart relation -> XSSFChart` directly instead
  of scanning every chart relation and trusting POI's optional `chart.getGraphicFrame()`
  backpointer. Mixed live-chart plus frameless-chart states now degrade cleanly instead of
  crashing factual drawing/chart inspection, explicit graphic-frame snapshotting no longer loses
  formula-resolution context when POI leaves the chart backpointer unset, the reproduced
  frameless-chart relation crash is now promoted as a committed success seed, and multi-chart
  copied sheets are now covered by focused regressions so orphaned chart-frame associations
  cannot drift back in quietly after the `0.51.0` release.
- Excel-facing sheet layout limits are now consistent across the contract, engine, CLI catalog,
  and docs. Row height now enforces Excel's real `409.0`-point ceiling instead of POI's wider
  twip storage envelope, `SET_SHEET_PRESENTATION.sheetDefaults` now validates the same default
  row/column limits as explicit sizing commands, `SET_SHEET_ZOOM` is now registered in the public
  limitations registry, and malformed drawing packages with empty embedded-object bytes now
  degrade into truthful zero-byte readback instead of crashing workbook inspection or Jazzer
  round-trip verification.
- The release protocol now covers the real bootstrap case where the primary checkout holds the
  unpublished release payload but release verification must run from a clean worktree. The
  documented flow now requires an explicit bootstrap branch or exported patch before applying that
  payload in the release worktree, so a dirty or problematic primary checkout no longer forces
  operators into an undocumented release path.

## [0.51.0] - 2026-04-22

### Fixed

- Sheet copy now repairs embedded-object sheet relationships that Apache POI XSSF `cloneSheet`
  leaves behind for OLE package parts, so copied embedded objects survive in-memory inspection,
  save/reopen, and the reproduced Jazzer `COPY_SHEET` round-trip regression. The failing fuzz case
  is now promoted as a committed success seed.
- The release closeout protocol is now executable instead of prose-only. GridGrind now ships
  `scripts/verify-release-primary-checkout.sh`, a dedicated shell regression for it, and updated
  release docs/check wiring so releasing from a disposable worktree cannot quietly leave the
  primary checkout behind `origin/main` with stale version-bearing files and misleading overlays.
- The supported Jazzer wrapper surface is now part of the tracked repository instead of hidden
  behind ignored local `jazzer/bin/` files. Clean release worktrees now carry the same
  `jazzer/bin/*` operator scripts as the primary checkout, and a new regression fails if the
  documented Jazzer wrapper surface drifts from what a clean checkout actually ships.

## [0.50.0] - 2026-04-22

### Fixed

- GridGrind now exposes Apache POI/XSSF workbook custom-XML mapping workflows as first-class
  public contract operations. `IMPORT_CUSTOM_XML_MAPPING`, `GET_CUSTOM_XML_MAPPINGS`, and
  `EXPORT_CUSTOM_XML_MAPPING` are now wired end to end across the contract, engine, executor,
  shipped examples, and public docs, with source-backed XML import support and focused regression
  coverage.
- Dedicated array-formula workflows are now documented truthfully across the public surface.
  README, reference docs, and the POI/XSSF capability inventory now distinguish scalar
  `FORMULA` cell writes from `SET_ARRAY_FORMULA` / `CLEAR_ARRAY_FORMULA` / `GET_ARRAY_FORMULAS`
  instead of incorrectly describing array formulas as wholly unsupported.
- The Apache POI/XSSF public docs now match the audited Apache POI 5.5.1 surface more
  truthfully. `docs/POI_EXCEL_CAPABILITY_INVENTORY.md` no longer overstates upstream support for
  sparklines or threaded comments, now splits custom XML mappings from slicers, and now lists
  previously omitted non-productized families such as signature lines and broader XDDF chart
  authoring. `docs/LIMITATIONS.md` now states that row/column worksheet bounds are already
  enforced on relevant request paths, and it no longer implies that Apache POI XSSF can only
  handle plain `.xlsx` files. Contract-side doc and POI-runtime audit tests now fail the build if
  these claims drift again or if the upstream jar surface changes underneath the published
  inventory.
- Sheet copy now rides POI XSSF's native clone path for the drawing family and then reapplies
  GridGrind-owned repair passes for workbook-core details such as formulas, raw data validations,
  conditional formatting, tables, sheet-owned autofilters, local names, comments, print layout,
  and protection metadata. Copied sheets now preserve supported pictures and charts instead of
  silently dropping them.
- Multi-plot combo-chart authoring is now verified directly in the engine test suite, and the
  public POI/XSSF capability inventory and operations docs now describe combo charts as a shipped
  supported capability instead of an undocumented gap.
- Signature-line drawing metadata is now exposed as a first-class `.xlsx` surface. GridGrind now
  ships `SET_SIGNATURE_LINE`, factual `SIGNATURE_LINE` drawing-object readback, authored anchor
  replacement and delete support for named signature lines, a generated signature-line example,
  and public docs that describe the real XSSF picture-format surface (`GIF`, `TIFF`, `EPS`,
  `BMP`, and `WPG` included) instead of the old narrower list.
- Release-surface shell verification is now resilient on this macOS baseline too. The packaged
  CLI/publication verifiers and their regression scripts now use repo-local disposable scratch
  under `tmp/` instead of brittle `/var/folders` temp allocation, and the release merge-handoff
  plus candidate-tag regression harnesses now fake the remote Git fetch surface directly instead
  of depending on flaky local clone/push transport behavior.
- `--print-protocol-catalog --operation` now exposes the nested and plain type groups that
  operators actually need for black-box request authoring. Type-group lookups such as
  `nestedTypes:cellInputTypes` and `plainTypes:chartInputType` now work directly instead of
  forcing callers to download and parse the full catalog just to discover cell, chart, or source
  payload shapes.
- Formula-backed chart references now evaluate authoritatively during chart authoring and
  readback instead of trusting stale worksheet or OOXML chart caches. Reference-backed series and
  formula titles now round-trip deterministically through save/reopen and the reproduced Jazzer
  crash is promoted as a committed success regression seed.
- Root `./check.sh` stall diagnostics now own diagnostic subprocess lifecycles end to end. The
  timeout path was extracted into a dedicated process-support helper, `capture_with_timeout` now
  escalates from `TERM` to `KILL` across the full captured process tree, and a shell regression
  now proves that TERM-ignoring parent/child trees do not leave orphaned `jcmd`-style processes
  behind.

## [0.49.0] - 2026-04-21

### Changed

- GridGrind now publishes a contract-owned intent-discovery layer on top of the exact protocol:
  `--print-task-catalog`, `--print-task-plan <id>`, `--print-task-keyword-match "<goal>"`, and
  `--doctor-request`. The packaged JAR and Docker artifact verifiers now black-box these task,
  planning, and diagnostics surfaces too, so the intent layer cannot drift silently from the core
  protocol contract.

### Fixed

- The thin CLI and packaged artifact help surface now behave truthfully for black-box operators:
  invoking GridGrind with no arguments and empty stdin prints help instead of falling through into
  a JSON parse failure, and `--print-protocol-catalog --operation <id>` now rejects ambiguous raw
  ids explicitly while teaching the canonical `<group>:<id>` lookup form.
- Goal-plan normalization now handles common plural stems more truthfully for freeform task
  discovery, including `quizzes -> quiz` and `sizes -> size`, instead of degrading ranked task
  matches with avoidable stemming drift.
- Chart readback now prefers live worksheet values over stale embedded chart caches for
  reference-backed series, so packaged artifact smoke and operator readback reflect the workbook's
  current sheet state instead of outdated OOXML cache text.
- Failure `context.stage` is now truthful and round-trippable for `RESOLVE_INPUTS`,
  `CALCULATION_PREFLIGHT`, and `CALCULATION_EXECUTION`; the calculation stage discriminator is now
  owned by concrete context types instead of a mismatched dynamic field.
- Synthetic CLI failure journals no longer invent `"unknown-plan"` or `"UNKNOWN"` source and
  persistence types before a request has parsed. Pre-parse failures now omit unavailable plan and
  transport facts instead of serializing misleading sentinel values.
- Step-target parsing now reports disallowed selector types truthfully, including ambiguous
  selector-family overlaps such as `BY_NAME`, instead of swallowing earlier parse failures and
  blaming the last attempted selector shape.
- `WorkbookPlan.steps()` now stays immutable after construction, so callers can no longer mutate a
  parsed or authored plan in place and bypass the duplicate-`stepId` invariant.
- Request-stream transport is now bounded at 16 MiB. Oversized JSON requests fail early with a
  product-owned `INVALID_REQUEST` message, and the streamed request/response/catalog write APIs no
  longer build whole-payload byte arrays before writing to caller-owned output streams.
- The machine-readable catalog and CLI help now publish step-target selector families directly.
  `mutationActionTypes`, `assertionTypes`, and `inspectionQueryTypes` now expose
  `targetSelectors` and `targetSelectorRule`, table and pivot mutations now declare only
  `BY_NAME_ON_SHEET`, and ambiguous shared selector ids such as `BY_NAME` are rejected explicitly
  instead of being guessed by parser order.
- `WorkbookStepJsonDeserializer` now uses Jackson 3's non-deprecated string-node APIs, and
  print-layout margin reads and writes now use Apache POI's `PageMargin` enum surface instead of
  the deprecated short-constant overloads. `ExcelStreamingWorkbookWriter` no longer carries a
  deprecated no-op `Cell.setCellType(FORMULA)` call after `setCellFormula(...)`, and its
  `close()` path now relies on `SXSSFWorkbook.close()`, which already disposes SXSSF temp files in
  Apache POI 5.5.1.
- `GridGrindJson` no longer emits `AlmostJavadoc` warnings from package-private parser-message
  helpers. The production JSON wording helpers now use real Javadoc comments so static analysis
  output stays signal-rich.
- Stale promoted Jazzer metadata for invalid protocol-request seeds is now refreshed against the
  current selector contract. The refresh test fixture now records the modern
  `EXPECTED_INVALID`/`INVALID_REQUEST_SHAPE` replay surface instead of preserving pre-hard-break
  authored source facts after parsing has already failed.
- Encrypted `.xlsx` package handling now models unencrypted versus password-protected open state
  explicitly instead of relying on nullable password sentinels, and POI relation-removal failures
  no longer mask fatal JVM errors as ordinary state exceptions.
- Exact-cell read validation is now single-sourced end to end. `GET_CELLS` address lists now share
  one canonical duplicate, bounds, and shape validator across the contract and engine while still
  preserving the indexed request messages callers already rely on.
- The Java workbook engine now exposes formula evaluation, capability inspection, cache clearing,
  and recalc-on-open toggles through a dedicated `workbook.formulas()` surface instead of crowding
  those operations onto the root workbook wrapper, and distinct gradient fills now keep their own
  OOXML entries even when Apache POI's public equality checks would otherwise alias different
  gradient geometries together.
- Executor and engine internals are now split into narrower helpers for execution validation,
  request path facts, calculation workflows, response shaping, sheet annotations, prepared chart
  models, chart source and snapshot handling, drawing removal and binary cleanup, drawing-chart
  flows, and drawing anchor state. The public contract stays the same, but the request executor
  and workbook engine are less monolithic and easier to evolve safely.
- Canonical operation, assertion, and inspection ids, execution-mode limits, CLI help labels, and
  shared discovery/error wording now derive from contract-owned structured metadata instead of
  parallel string copies. The CLI help surface is now a typed catalog section tree, execution-mode
  validation messages come from one shared metadata owner, and the build now fails if public
  docs/help/catalog/example surfaces mention an unregistered canonical mutation, assertion, or
  inspection id.

## [0.48.0] - 2026-04-19

### Changed

- GridGrind no longer ships a monolithic `protocol` module. The public contract, metadata
  registry, and JSON codecs now live in `contract`, while `executor` is the only execution bridge
  into the workbook engine. This also puts the Java-first contract replacement program into formal
  hard-break mode.
- The nested Jazzer build now consumes the `executor` plus `contract` split instead of the deleted
  monolithic `protocol` module, and the `contract` module's 100 % coverage gate now includes the
  downstream `executor` and `cli` consumer tests that exercise the canonical public contract after
  the split.
- The public contract is now selector-first end to end. Top-level coordinate and sheet-name
  fields have been removed from request operations and reads in favor of canonical selector
  payloads, and the public README, operations reference, quick reference, shipped examples, and
  release gates now reject the deleted pre-selector vocabulary.
- The public request model is now step-only end to end. The deleted `operations[]` plus `reads[]`
  split has been replaced by ordered `steps[]`, step envelopes no longer carry a redundant outer
  discriminator, and public request/response docs, Docker smoke, replay fixtures, and contract
  surface checks now reject the deleted `selector` and `requestId` vocabulary alongside the old
  arrays.
- GridGrind now ships first-class assertion steps in the public contract. Ordered plans can verify
  workbook state with canonical `ASSERTION` steps, success responses return ordered
  `assertions[]`, and public docs, examples, CLI help, and artifact-surface verification now
  describe the mutate-then-verify model instead of the older mutation-plus-read-only framing.
- GridGrind responses now always carry a structured execution `journal`, and the top-level request
  execution policy is now the canonical `execution` envelope with nested `mode` and `journal`
  configuration instead of the older top-level `executionMode` sidecar. `VERBOSE` journal mode is
  now a public contract feature that streams fine-grained live execution events to CLI stderr
  while preserving the same events in the response payload.
- GridGrind no longer exposes formula evaluation, cache clearing, or recalc-on-open as mutation
  step types. Those deleted actions are replaced with the canonical top-level
  `execution.calculation` policy, whose shipped strategies are `EVALUATE_ALL`,
  `EVALUATE_TARGETS`, `CLEAR_CACHES_ONLY`, and `DO_NOT_CALCULATE` plus optional
  `markRecalculateOnOpen=true`.
- GridGrind mutation payloads are now source-backed end to end. Text-bearing authored fields use
  canonical `INLINE`, `UTF8_FILE`, or `STANDARD_INPUT` sources, binary-bearing authored fields use
  canonical `INLINE_BASE64`, `FILE`, or `STANDARD_INPUT` sources, and the structured execution
  journal now records authored input loading under `journal.inputResolution` before workbook open.
- GridGrind now ships the `authoring-java` module as a first-class fluent Java surface over the
  canonical contract. Java callers can build selector-first mutation, inspection, assertion, and
  execution-policy plans without hand-writing JSON. The authored target records now reject null
  selectors immediately instead of carrying invalid state until later execution, and the shipped
  `examples/java-authoring-workflow.java` example is compile-verified against that published API.
- CLI help, built-in examples, and the machine-readable catalog are now contract-owned surfaces.
  The thin CLI transport renders help from the `contract` module, `--print-example <id>` emits one
  generated built-in example request, `--print-protocol-catalog` publishes `cliSurface` plus
  `shippedExamples`, and the committed `examples/*.json` fixtures are regenerated from that same
  registry instead of being hand-maintained JSON.
- The root `check` gate now includes explicit-import verification for handwritten production Java
  and Kotlin sources, so wildcard imports fail the canonical build instead of relying on reviewer
  cleanup.
- Release and fuzz hardening now treat packaged artifacts as the authoritative contract surface.
  CLI contract verification reads structured help plus shipped-example lines from the built JAR or
  Docker image, Docker smoke rebuilds the CLI fat JAR before packaging so it cannot validate stale
  output, and promoted Jazzer metadata refresh rewrites dead source pointers to the live promoted
  inputs when regenerated examples retire old fixture paths.
- Named-range deletion now accepts only exact scoped selectors in the canonical Java contract.
  Workbook-scoped and sheet-scoped named ranges remain the same JSON wire types, but the hard-
  break Java surface no longer allows broad named-range selector families where authoritative
  deletion semantics require an exact scope.
- The committed `protocol-workflow` Jazzer corpus now uses neutral `workflow_case_##` seed names
  instead of semantic blob names. Those inputs are opaque generator bytes, so the authoritative
  decoded behavior now lives only in refreshed replay metadata and replay text.
- Jazzer replay metadata and replay text for protocol request/workflow harnesses now record
  assertion counts and assertion kinds explicitly, so assertion-bearing opaque workflow seeds no
  longer look like unexplained response-kind flips during regression replay.
- The release protocol now treats open Dependabot PRs as first-class release hygiene. Release-time
  pre-flight now requires explicitly identifying open Dependabot work, and after the public
  release is verified each Dependabot PR must be merged, closed, or consciously kept open with a
  stated reason; stale automation branches are no longer acceptable release leftovers.
- Workbook-protection readbacks, docs, and parity checks now use the same canonical field name
  `revisionsLocked` end to end instead of mixing singular and plural variants across write and
  read surfaces.
- Formula-backed chart titles are now authored with an explicit OOXML string cache, so numeric
  title cells survive `.xlsx` save/reopen exactly instead of drifting to stale cached values
  during chart round-trips.

### Fixed

- Chart reads now resolve formula-backed title `cachedText` from the referenced cell when OOXML
  omits a cached string, so authored chart titles no longer come back blank while the same chart's
  formula still points at a real text cell.
- Existing bar, line, and pie charts now treat `Title.None` series updates idempotently, so
  removing a series title no longer falls into Apache POI `unsetTx()` crashes when the underlying
  chart XML never carried a `<c:tx>` node.
- The committed Jazzer regression floor now includes the `fractional_integer_field.json` invalid-
  request seed with matching promotion metadata and replay text, so the nested fuzz build no
  longer drifts when integer-field shape validation expands.
- The nested Jazzer replay floor now refreshes both selector-first protocol-request fixtures and
  neutralized protocol-workflow binary metadata against the current replay engine, so contract
  topology changes no longer leave stale replay expectations behind.
- The Docker smoke gate now verifies low-memory streaming in the same two-step shape the product
  actually supports: `STREAMING_WRITE` authoring first, then summary-only `EVENT_READ` readback
  against the materialized workbook.
- Table-aware exact-cell selectors now execute truthfully end to end. Selector-first table-key
  targeting resolves through the canonical executor for mutations, inspections, and assertions,
  source-backed row-key values now load before execution just like source-backed mutation payloads,
  duplicate-key matches fail instead of guessing, and zero-match inspections no longer collapse
  into lower-level table lookup failures.
- The Docker smoke and publication-surface regression gates now author `APPEND_ROW` text values
  with canonical source-backed payloads.
- Streaming-write calculation failures now surface as structured `CALCULATION_EXECUTION` contract
  failures before persistence instead of collapsing into generic request-level runtime handling.
- Drawing-object picture reads now report factual raster dimensions when the image format exposes
  them, and expanded row grouping no longer destroys pre-existing manual hidden-row state during
  group or ungroup operations.
- Protocol JSON parsing now rejects floating-point JSON numbers for integer contract fields instead
  of silently truncating them during deserialization.
- Selector and operation validation now preserve indexed null and invalid-entry diagnostics instead
  of letting early collection copying collapse those failures into opaque bare `NullPointerException`
  paths.
- The workbook-core inspection pipeline now uses `stepId` internally as well as on the public
  contract, so command/result plumbing, executor conversion, parity helpers, and inspection tests
  no longer translate back through the deleted `requestId` terminology.
- The nested Jazzer workflow generator, invariants, telemetry, and replay-safe support layer now
  exercise assertion steps as first-class workflow elements instead of fuzzing only mutations plus
  inspections, and the shipped `examples/assertion-request.json` fixture is regression-tested as a
  public mutate-then-verify request.
- Presence-style assertions now own selector-count semantics authoritatively: exact named-range
  and chart misses are evaluated as zero observed entities for
  `EXPECT_NAMED_RANGE_PRESENT` / `EXPECT_NAMED_RANGE_ABSENT` and
  `EXPECT_CHART_PRESENT` / `EXPECT_CHART_ABSENT` instead of leaking lower-level not-found
  failures, and the executor regression suite now locks that behavior in.
- Every response surface now exposes execution journaling consistently: success and failure payloads
  include structured phase and step telemetry, the shipped assertion example demonstrates
  `execution.journal.level=VERBOSE`, and Docker smoke now black-boxes both response-journal
  emission and live stderr event streaming from the packaged artifact.
- Source-backed authored input failures now surface as explicit `INPUT_SOURCE_NOT_FOUND`,
  `INPUT_SOURCE_UNAVAILABLE`, or `INPUT_SOURCE_IO_ERROR` contract failures instead of collapsing
  into generic request or filesystem errors, and the shipped
  `examples/source-backed-input-request.json` fixture plus promoted Jazzer replay floor now lock
  file-backed text, file-backed formulas, and file-backed binary payloads into regression coverage.
- Engine selection and read payload helpers no longer perform wasteful double-freeze list copies
  when validating already-immutable `List.copyOf(...)` inputs.
- The build and developer docs now spell out the upstream Jackson 3 rule correctly: Jackson 3
  databind intentionally still uses the `com.fasterxml.jackson.annotation` namespace, so GridGrind
  now guards that fact explicitly instead of making it look like accidental dependency drift.

## [0.47.0] - 2026-04-16

### Changed

- Public streaming-write docs, CLI help, the protocol catalog, and runtime validation messages now
  use the single canonical operation name `FORCE_FORMULA_RECALCULATION_ON_OPEN` consistently.
- Agent-facing catalog summaries now describe the actual response shape for `GET_SHEET_LAYOUT`,
  `GET_FORMULA_SURFACE`, `GET_NAMED_RANGE_SURFACE`, and the workbook-health analysis reads,
  including checked-count fields, `layout.presentation`, grouped formula summaries, and flat
  aggregated findings.
- Core-owned contract text now drives the thin downstream CLI help, protocol catalog discovery
  summaries, and execution-mode validation messages for the public low-memory and workbook-health
  surfaces, so operation-name and limitation wording no longer lives in multiple hand-maintained
  string copies.
- Release verification now treats `--help` and `--print-protocol-catalog` as first-class public
  artifacts. The build JAR, local Docker smoke image, and published GHCR tags are black-box
  checked for canonical streaming-write wording, the hard `LAMBDA`/`LET` boundary, and the richer
  analysis/layout catalog summaries before a release is considered healthy.

### Fixed

- Gradient-fill authoring now rejects mixed geometry models up front instead of serializing
  impossible `.xlsx` styles. `LINEAR` gradients accept `degree`, `PATH` gradients accept
  `left/right/top/bottom`, and valid gradient-plus-protection styles now survive `.xlsx`
  round-trips cleanly.
- Formula docs and CLI help now state the current hard limitation truthfully: authored
  array-formula braces are rejected as `INVALID_FORMULA`, `LAMBDA` and `LET` are currently
  rejected as `INVALID_FORMULA` because Apache POI cannot parse them, and loaded formulas that POI
  parses but cannot evaluate surface as `UNSUPPORTED_FORMULA`.
- Protocol JSON parsing no longer carries a special-case alias hint path for retired operation
  discriminator spellings; invalid discriminators now fail uniformly as unknown type values.
- The Apache POI XSSF capability inventory now cites the real comment implementation files instead
  of a stale nonexistent comment-support evidence path.

## [0.46.0] - 2026-04-15

### Changed

- The release protocol now requires an explicit post-merge `main` CI handoff before tagging, and
  `scripts/verify-release-merge-handoff.sh` plus its shell regression test now enforce that the
  checked-out release commit matches `origin/main` and already has green `Check` and `Docker smoke`
  runs before any public tag is created.
- The machine-readable protocol catalog and CLI help now describe the full
  `ANALYZE_WORKBOOK_FINDINGS` aggregate correctly, including pivot-table health, and the public
  workbook-health docs now point directly at the shipped batch-analysis request examples.

### Fixed

- The public Apache POI XSSF capability inventory no longer collapses distinct read and analysis
  operations into shared rows. It now names the individual formula, hyperlink, named-range, sheet
  schema, and aggregate workbook-finding surfaces explicitly, and it declares array formulas and
  sparklines as `NOT_EXPOSED` instead of leaving them implicit.
- Public formula docs now spell out the real request-contract boundary: authored formulas are
  scalar only, array-formula braces are rejected as `INVALID_FORMULA`, newer constructs such as
  `LAMBDA`/`LET` may fail during parse when Apache POI cannot read them, and loaded formulas that
  POI parses but cannot evaluate surface as `UNSUPPORTED_FORMULA`.
- The committed `examples/*.json` request fixtures are now regression-tested for JSON/schema
  validity so public examples cannot drift silently from the shipped protocol.

## [0.45.0] - 2026-04-14

### Changed

- Release and container workflow-dispatch reruns now refuse to publish unless the checked-out tag
  still matches `gradle.properties`, resolves to the exact remote tag commit, remains reachable
  from `main`, and already has successful `Check` plus `Docker smoke` CI runs on that commit.
- Published GHCR release images now rebuild from a digest-pinned Azul Java 26 base image and emit
  explicit OCI provenance plus SBOM attestations.
- The root `coverage` and `jacocoAggregatedReport` tasks now discover JaCoCo-enabled Java
  subprojects dynamically and aggregate every module `build/jacoco/*.exec` file instead of
  hardcoding today's module list.
- The `protocol` module no longer leaks `jackson-databind` onto consumer compile classpaths; the
  runtime dependency remains available where needed without widening the published Gradle API.

### Fixed

- The container publication workflow no longer overrides the image's SPDX license label with a
  stale MIT-only value; published OCI metadata now matches the shipped notices and licenses.
- `./check.sh` Stage 4 now regression-tests the release-candidate tag gate and publication
  workflow contract, so pinned-base-image, attestation, and publication-policy drift fails
  locally before a public release run.

## [0.44.0] - 2026-04-14

### Changed

- JaCoCo branch-coverage enforcement raised toward 100 % for the `engine` module: branch paths
  previously unreachable by the test suite are now exercised across seven controller and support
  classes (details below). Coverage targets are asserted by the Gradle `coverage` task and must
  remain green before any release.
- The remaining advanced-XSSF branch gaps are now closed across drawing, pivot-table,
  package-security, autofilter, print-layout, data-validation, sheet-state, event-read,
  sheet-copy, conditional-formatting, formula-rename, and formula-exception flows. Helpers that
  were only guarding unmaterializable POI or XmlBeans null states were simplified to match the
  real runtime contract, while new tests now exercise pivot-cache registration cleanup, embedded
  preview lookup, sparse event-reader workbook metadata, and package-signing or encryption edge
  handling.
- `ExcelPrintLayoutController` branch coverage extended: `shouldUnsetPageSetupOrientation` now
  covers the `LANDSCAPE` false-return path in addition to `PORTRAIT`; `isEmptyPageSetup` now
  covers all eleven individual attribute-present/default combinations (paperSize, draft,
  blackAndWhite, copies, useFirstPageNumber, firstPageNumber) for both the true and false returns.
- `ExcelConditionalFormattingController` health-check branch coverage extended: the
  `CellValueRule`-with-two-valid-formulas path (both `isBrokenFormula` calls return false, no
  health finding emitted) is now exercised by a dedicated test so the short-circuit `&&` chain
  cannot regress silently.
- `ExcelWorkbook` persistence-options branch coverage extended: the
  `persistenceOptions != null && persistenceOptions.isEmpty()` fast-path (treat as plain save)
  is now covered by a test that passes `new ExcelOoxmlPersistenceOptions(null, null)` to
  `save(Path, ExcelOoxmlPersistenceOptions)`.
- `ExcelFormulaSheetRenameSupport` external-workbook branch coverage extended: the
  `getExternalWorkbookNumber() >= 1` guard (skip rename for cross-workbook references) is now
  covered via a test that links an external workbook, references it from a formula, and asserts
  the returned formula still names the external book rather than the renamed local sheet.
- `FormulaExceptions.isKnownBuiltinFunction` branch coverage extended: the
  `AnalysisToolPak.getSupportedFunctionNames()` and `AnalysisToolPak.getNotSupportedFunctionNames()`
  OR-chain branches (conditions 4 and 5) are now covered by tests that locate functions present
  exclusively in each ATP set — absent from `FunctionMetadataRegistry` and both
  `FunctionEval` catalogs — so each OR branch can be the first-true entry.
- `ExcelSheetCopyController` validation-formula retargeting branch coverage extended: edge cases
  now tested include formulas shorter than two characters, quoted-list literals that start but do
  not end with `"`, absent `formula1`, blank `formula1`, absent `formula2`, and blank `formula2`,
  exercising all guard branches in `isQuotedListLiteral` and the `isSetFormula`/blank checks.
- `ExcelEventWorkbookReader` branch coverage extended: the `sizeOfWorkbookViewArray() == 0` path
  in `activeSheetIndex` is now covered; additionally an integration test exercises blank cell-ref
  (`r=""`) handling and a `<col>` element with no `max` attribute in the sheet XML parse path.
- README, the public capability inventory, the internal XSSF parity ledger, and the Jazzer
  coverage guide now agree on the shipped verification surface for drawings, charts, pivots,
  conditional formatting, low-memory event reads, and OOXML package security. The Jazzer guide no
  longer lists already covered chart, picture, pivot, or conditional-format families as missing
  fuzz surface.

### Fixed

- Docker smoke now builds through `docker buildx build --load` under an anonymous `DOCKER_CONFIG`
  while preserving the active local Docker engine endpoint, so local and CI verification no longer
  depend on personal Docker login state or Docker's deprecated legacy builder path.
- Docker smoke now runs mounted-path container commands as the caller's UID:GID, so generated
  response files and saved workbooks stay operator-owned instead of leaving root-owned artifacts
  behind on Linux CI or other Unix-like hosts.
- `scripts/verify-container-publication.sh`, the Docker developer docs, and the release protocol
  now codify anonymous public-container verification while preserving the active Docker engine
  endpoint, keeping release verification aligned with the real workstation contract.
- `NOTICE` now covers all bundled runtime dependencies: Apache Santuario xmlsec (Apache 2.0),
  Bouncy Castle bcpkix/bcprov/bcutil (MIT variant), and SLF4J API (MIT). These were added as
  runtime dependencies in 0.43.0 but omitted from the attribution file.
- `--license` output now includes the `NOTICE` attribution file between the GridGrind MIT license
  and the third-party license texts, satisfying Apache License 2.0 §4(d) for runtime distributions.
- `PATENTS.md` dependency table now lists Apache Santuario xmlsec, Bouncy Castle, and SLF4J API.
- Fat JAR (`META-INF/`) now includes `PATENTS.md` alongside `NOTICE` and the license files.
- Docker image now ships `PATENTS.md` in `/usr/share/doc/gridgrind/` for completeness.
- `org.opencontainers.image.licenses` OCI label updated from `MIT` to the full SPDX expression
  `MIT AND Apache-2.0 AND BSD-3-Clause` to accurately reflect the bundled dependency licenses.
- `scripts/verify-container-publication.sh` now verifies the exact two-line `--version` product
  header that the shipped container exposes (`GridGrind <version>` plus the product
  description), and `./check.sh` now includes a dedicated shell regression for that verifier so
  future release workflows cannot fail on a stale output assumption.
- The packaged CLI and Docker image now include the Bouncy Castle runtime dependencies required by
  OOXML package-security inspection, so `source.type=EXISTING` opens and low-memory
  `STREAMING_WRITE` readback no longer fail at runtime with `NoClassDefFoundError` on the
  existing-workbook path.
- The packaged CLI now also bundles an explicit SLF4J provider for package-security dependencies,
  and the Docker smoke gate rejects unexpected stderr on request execution, so successful
  existing-workbook and low-memory request flows stay clean for agent consumers.
- The Jazzer runtime now binds the same SLF4J provider explicitly and asserts it in support
  tests, so regression replay and round-trip harnesses no longer emit fallback logger warnings
  on stderr while they pass.
- The machine-readable protocol catalog now correctly marks `SET_SHEET_PROTECTION.password` and
  `SET_AUTOFILTER.criteria` or `sortState` as optional to match the shipped runtime contract.
- Public request-shape docs and catalog summaries now clarify sparse append-edge row or column
  inserts: valid append-position inserts do not materialize a new physical tail row or column
  until content or explicit metadata exists there.

## [0.43.0] - 2026-04-13

### Added

- `--license` flag: prints the GridGrind MIT license followed by third-party dependency license
  notices (Apache 2.0, BSD-3-Clause). License texts are bundled in the JAR under
  `licenses/` and read at runtime.
- Added `SET_SHEET_PRESENTATION` plus `GET_SHEET_LAYOUT.presentation` for sheet display flags,
  right-to-left layout, tab color, outline-summary placement, default row or column sizing, and
  ignored-error suppression.
- Added `printGridlines` to print-setup authoring and factual `GET_PRINT_LAYOUT` readback.

### Changed

- `--version` now prints the same two-line product header as the opening of `--help`
  (`GridGrind <version>` on the first line, the product description on the second), providing
  a single source of truth for product identity output via `productHeader()`.
- `--help` banner changed from `GridGrind CLI <version>` to `GridGrind <version>` to match
  the product name.
- Promoted the Apache POI XSSF capability inventory into a public reference at
  [docs/POI_EXCEL_CAPABILITY_INVENTORY.md](./docs/POI_EXCEL_CAPABILITY_INVENTORY.md) and rewrote
  it around the shipped `.xlsx` contract instead of internal parity-phase language.
- README now links directly to the public capability inventory alongside the operations,
  quick-reference, and limitations docs.
- Public docs, quick-reference snippets, and runnable structural-layout examples now describe the
  shipped sheet-presentation surface and print-gridline output explicitly, and the README example
  guide no longer uses internal phase labels.

### Fixed

- The parity-doc regression test now verifies the public capability inventory path, release
  version, public-facing wording, and the absence of the retired internal parity-planning files.
- The POI capability inventory no longer underreports the shipped sheet-view and print-layout
  surface; it now reflects the public contract accurately.
- Sheet-presentation and print-setup model validation no longer carry dead compatibility overloads
  or unreachable null-element branches, so the Java contract stays smaller and the regression floor
  asserts only live behavior.

## [0.42.0] - 2026-04-13

### Added

- Added OOXML package-security support to the public `.xlsx` contract:
  `source.security.password` for encrypted existing sources, `persistence.security.encryption`,
  `persistence.security.signature`, and factual `GET_PACKAGE_SECURITY` readback for package
  encryption and package-signature state.
- Added `examples/package-security-create-request.json`
  and [examples/package-security-inspect-request.json](./examples/package-security-inspect-request.json),
  a paired public example flow for encrypted workbook authoring followed by factual package-security
  inspection.
- Added a promoted Jazzer protocol-request seed for the package-security request surface so
  `source.security`, `persistence.security`, and `GET_PACKAGE_SECURITY` stay replay-verified in the
  committed regression floor.

### Changed

- Public docs, quick-reference snippets, README guidance now describe the shipped OOXML
  package-security surface explicitly instead of treating encryption and signing as absent.

### Fixed

- Legacy OLE2 `.xls` files are no longer misclassified as encrypted OOXML packages on the
  package-security open path; unsupported legacy workbooks now fail honestly as non-`.xlsx`
  inputs instead of incorrectly demanding `source.security.password`.
- The parity ledger now verifies actual encrypted-open, encrypted-save, signed-read,
  signed-authoring, invalid-password, and invalid-signature behavior against the committed corpus
  instead of keeping legacy "gap" probes after the runtime support landed.
- The invalid-signature parity corpus now tampers signed workbooks through a separate output path
  instead of rewriting the signed package in place, so the corpus materializes as genuinely
  `INVALID` rather than as a broken ZIP stream.

## [0.41.0] - 2026-04-13

### Added

- Added top-level `executionMode` request support for low-memory `.xlsx` workflows:
  `readMode: EVENT_READ` for summary-only event-model reads and `writeMode: STREAMING_WRITE` for
  append-oriented SXSSF authoring on `NEW` workbooks.
- Added [examples/large-file-modes-request.json](./examples/large-file-modes-request.json), a
  runnable example covering `executionMode`, `STREAMING_WRITE`, and summary-only `EVENT_READ`
  readback.

### Changed

- Public docs, CLI help, the limitations registry now describe the shipped low-memory execution
  contract explicitly instead of treating event reads and streaming writes as absent.

### Fixed

- `ExecutionModeInput.isDefault()` is now JSON-ignored, so request round-trips no longer leak a
  stray `executionMode.default` field.
- The streaming writer now owns row advancement explicitly and disposes SXSSF temp files on close
  instead of relying on incidental POI state.
- Parity probes now verify actual low-memory request behavior against the large-sheet
  corpus instead of only checking for missing catalog placeholders.
- Jazzer replay support now resolves promoted large-file example inputs correctly from both the
  `jazzer` module root and the repository root, so the committed low-memory example stays
  replay-verified under the real Gradle execution layout.

## [0.40.0] - 2026-04-13

### Added

- Added limited XSSF pivot-table parity to the public `.xlsx` contract with `SET_PIVOT_TABLE`,
  `DELETE_PIVOT_TABLE`, `GET_PIVOT_TABLES`, and `ANALYZE_PIVOT_TABLE_HEALTH`.
- Added [examples/pivot-request.json](./examples/pivot-request.json), a runnable example covering
  range-backed, named-range-backed, and table-backed pivot authoring plus factual pivot readback
  and pivot-health analysis.
- Added the matching promoted Jazzer protocol-request seed so the public pivot example is
  replay-verified in the committed regression floor.

### Changed

- Public docs, quick-reference snippets, README guidance, now describe the shipped pivot-table
  surface explicitly, including supported source kinds, authored anchor rules, explicit
  unsupported readback, and pivot-health analysis.

### Fixed

- Pivot-table authoring and preservation now normalize named-range sources correctly, keep
  workbook-global pivot names and cache relations stable across save or reopen, and rebuild
  workbook pivot wiring after create or delete flows instead of depending on stale POI relation
  allocation state.
- Replacing a pivot table in place on the same sheet now cleans up orphaned cache-record parts and
  re-primes POI's pivot-part allocator from the workbook package, so repeated authoring no longer
  collides on stale `/xl/pivotCache/pivotCacheRecords*.xml` numbering after delete or reopen flows.
- Malformed or oversized pivot number-format identifiers now degrade into truthful readback instead
  of throwing during factual snapshotting or parity-oracle reporting.
- Jazzer request labeling, workflow invariants, workbook-shape checks, `.xlsx` round-trip
  verification, replay expectations, and promoted workflow metadata now model the current
  pivot-table and protocol-workflow contract directly instead of lagging behind the shipped
  behavior.

## [0.39.0] - 2026-04-12

### Added

- Added [examples/chart-request.json](./examples/chart-request.json), a runnable example covering
  supported `BAR` chart authoring, named-range-backed series binding, explicit chart-anchor
  replacement, factual `GET_CHARTS` readback, and matching chart inventory in
  `GET_DRAWING_OBJECTS`.
- Added the matching promoted Jazzer protocol-request seed plus deterministic Jazzer support
  coverage for chart authoring and chart readback, so the public example is replay-verified and
  the fuzz-support layer now exercises the chart contract directly.

### Changed

- Public docs, quick-reference snippets, and README guidance now document the shipped chart
  contract explicitly: `SET_CHART`, `GET_CHARTS`, supported simple `BAR` or `LINE` or `PIE`
  families, named-range-backed series formulas, and explicit `UNSUPPORTED` readback for
  unsupported plot families.
- `SET_SHAPE` and `SET_CHART` validation is now explicitly non-mutating: failed authored shape or
  chart requests leave existing drawing state untouched instead of leaking partial artifacts.

### Fixed

- Jazzer `.xlsx` round-trip verification now asserts chart and drawing-object preservation across
  reopen instead of treating charts as an untracked blind spot.
- Jazzer request labeling, workflow-shape validation, and workbook-shape invariants now model
  `SET_CHART`, `GET_CHARTS`, and chart-backed drawing inventory as first-class protocol surface.
- Failed `SET_SHAPE` preset validation and failed `SET_CHART` preflight no longer leave partial
  shapes, chart frames, or half-mutated existing charts behind.
- Chart factual reads now normalize blank stored OOXML titles to `NONE`, preserve sparse literal
  cache positions as empty-string gaps instead of aborting the read, and degrade broken chart
  relationships into truthful surviving drawing facts when a graphic frame remains.
- The chart controller now owns explicit POI translation and relation-removal seams, so
  chart-family enum mapping and chart-part deletion are regression-tested directly instead of
  hiding inside one large controller branch.

## [0.38.0] - 2026-04-12

### Added

- Added `examples/drawing-media-request.json`, a runnable
  example covering picture, shape, and embedded-object authoring, explicit drawing-anchor
  replacement, drawing payload extraction, and comment coexistence on the same sheet.
- Added the matching promoted Jazzer protocol-request seed plus deterministic Jazzer support
  coverage for drawing-media workflows, so the public example is replay-verified and the nested
  round-trip or invariant layer now asserts the drawing contract directly.

### Changed

- Public docs, quick-reference snippets, README guidance, and the internal XSSF parity and
  inventory records now describe the shipped drawing, image, and embedded-object platform
  explicitly, including authored two-cell anchors, factual read-side anchor variants, and
  drawing-payload extraction boundaries.

### Fixed

- Jazzer sequence labeling, workflow generation, response invariants, and `.xlsx` round-trip
  verification now model the drawing-media command and read surface instead of silently lagging
  behind it.
- The direct-POI parity oracle now stores embedded-object payload bytes behind a defensive
  immutable class, keeping the parity build warning-free under Error Prone's array-component
  checks.

## [0.37.0] - 2026-04-12

### Added

- Added `examples/formula-environment-request.json`,
  a runnable example covering top-level `formulaEnvironment`, template-backed UDF
  registration, targeted formula evaluation, and explicit formula-cache clearing.
- Added the matching promoted Jazzer protocol-request seed for the formula-environment example, so
  the public request now stays replay-verified in the committed regression floor.

### Changed

- Public docs, quick-reference snippets, README guidance, and the internal XSSF parity records now
  describe the completed formula-evaluation contract explicitly: external workbook bindings,
  missing-workbook policy control, template-backed UDF toolpacks, targeted formula evaluation, and
  explicit persisted-cache clearing.

### Fixed

- `CLEAR_FORMULA_CACHES` now clears persisted formula cached results in the workbook while normal
  post-mutation invalidation still resets only the in-process evaluator cache, so explicit
  lifecycle control is honest without changing ordinary mutation parity semantics.
- The direct-POI parity oracle now evaluates external-link and UDF scenarios through read-only
  input streams before applying transient evaluator configuration, so parity measurement no longer
  mutates the corpus workbook it is measuring.
- Jazzer generation, labeling, and `.xlsx` round-trip verification now cover targeted formula
  evaluation and explicit cache clearing, so the fuzz support layer no longer lags behind the
  shipped formula contract.
- Formula-health analysis and request/command dispatch now route the completed formula-lifecycle
  families through explicit type-owned paths, removing stale unreachable fallback branches and
  making the verification surface match the runtime architecture more directly.

## [0.36.0] - 2026-04-11

### Added

- Added `examples/advanced-mutation-request.json`, a
  runnable workbook-core mutation example covering password-bearing protection, formula-defined
  named ranges, advanced table and autofilter mutation, advanced conditional formatting, rich
  comments, advanced page setup, and structured style colors.

### Changed

- Public docs, quick-reference snippets, README guidance, the internal XSSF capability inventory,
  and the internal parity execution spec now describe the completed non-drawing workbook-core
  mutation contract explicitly instead of the earlier partial summaries.
- `examples/advanced-readback-request.json` now
  materializes the richer factual readback surface it advertises, including workbook protection,
  rich comment runs and anchors, advanced page setup, structured style colors and gradients,
  autofilter criteria and sort state, and advanced table metadata.

### Fixed

- `SET_PRINT_LAYOUT` docs now cover the supported advanced page-setup payload instead of only the
  earlier core print-layout subset.
- Public contract docs now describe the real `SET_SHEET_PROTECTION`,
  `SET_WORKBOOK_PROTECTION`, `SET_COMMENT`, `APPLY_STYLE`, `SET_AUTOFILTER`, `SET_TABLE`,
  `SET_CONDITIONAL_FORMATTING`, and `SET_NAMED_RANGE` surfaces, including password-bearing
  protection, rich comments, structured color writes, gradient fills, advanced filter metadata,
  advanced table metadata, six conditional-format rule families, and formula-defined names.
- The internal XSSF parity oracle now detects workbook and revisions password-hash presence across
  both legacy and modern OOXML workbook-protection fields, so parity verification no longer reports
  a false regression on SHA-512-authored workbooks.

## [0.35.0] - 2026-04-11

### Added

- GridGrind now exposes read-parity surface for workbook protection, rich comment runs
  and anchors, advanced print setup, structured theme or indexed or tinted color facts,
  gradient fills, autofilter criteria and sort state, advanced table metadata, and the
  remaining POI-readable XSSF conditional-formatting families modeled by GridGrind.
- Added `examples/advanced-readback-request.json` plus
  the matching promoted Jazzer seed so the richer factual readback contract is both publicly
  demonstrated and replay-verified.

### Changed

- The CLI help now states that column structural edits are blocked by any workbook formulas, not
  just formulas on the sheet being edited.
- The protocol catalog, public docs, README, and the promoted `table_autofilter_request.json`
  Jazzer seed now use the real defaulted `SET_TABLE` contract: omit `showTotalsRow` unless the
  table actually includes a totals row.
- The public docs, quick reference, README, Jazzer docs, and parity records now describe the full
  richer readback contract instead of the earlier narrowed summaries.

### Fixed

- `SET_TABLE` now treats `showTotalsRow` as a genuinely optional request field that defaults to
  `false`, matching the protocol catalog and black-box CLI behavior.
- The protocol catalog builder now rejects any future attempt to mark a primitive record component
  as optional, so catalog optionality cannot drift away from JSON deserialization semantics again.
- Data-validation health analysis now preserves distinct malformed raw-validation states instead of
  collapsing them into broader findings, and row or column structural-edit guards now tolerate
  malformed raw validation records without crashing while computing unsupported-formula checks.
- Advanced readback now degrades malformed raw conditional-formatting family metadata and malformed
  raw data-validation enum metadata into factual unsupported reports instead of crashing before the
  workbook can be inspected.
- Persisted autofilter sort-state and sort-condition ranges are now reported exactly as stored,
  including blank raw ranges, so malformed workbook metadata is surfaced to callers instead of
  being rejected during factual readback.
- The advanced XSSF parity corpus now correctly materializes workbook autofilters, theme or tinted
  font colors, and gradient fills, so the parity oracle measures the intended read surface instead
  of an underspecified fixture subset.

## [0.34.0] - 2026-04-11

### Added

- Added an executable Apache POI `5.5.1` XSSF parity gate. GridGrind now ships a canonical parity
  ledger, a golden `.xlsx` corpus, a direct-POI oracle harness, GridGrind-side comparator probes,
  and a root `./gradlew parity` task for measuring current `.xlsx` parity status end to end.

### Changed

- Root and protocol `check` verification now includes the XSSF parity source set plus parity PMD
  coverage, so `.xlsx` parity drift fails local verification instead of living only in ad hoc
  investigation.
- Jazzer operator docs now distinguish the supported `jazzer/bin/*` surface from raw
  `./gradlew --project-dir jazzer ...` debugging more explicitly, and the seed-inventory docs now
  point at one authoritative exhaustive committed-input list instead of drifting count copies.
- Jazzer operator docs now state a single supported active-fuzz method only: `jazzer/bin/*`.
  Raw Gradle is documented only for deterministic nested-build verification, not as an endorsed
  alternative fuzz entrypoint.

### Fixed

- Style snapshot extraction no longer preserves border colors when the effective border style is
  `NONE`, avoiding impossible border-state reports for advanced POI-authored `.xlsx` workbooks.
- The parity corpus now materializes real agile-encrypted OOXML workbooks, and the direct-POI
  parity oracle opens them through POI's decryptor flow instead of treating them as plain `.xlsx`
  files.
- Nested Jazzer active fuzzing now preloads a project-owned premain agent that publishes startup
  instrumentation to Byte Buddy before Jazzer's JUnit extension runs, so Java 26 live fuzzing no
  longer wedges in the external attach path before the harness starts executing.
- Active Jazzer fuzzing now hard-fails on GitHub Actions, so GitHub remains a deterministic-only
  verification surface even if an active fuzz task is wired there by mistake.
- Promoted Jazzer metadata no longer carries stray committed `.txt.tmp` artifacts, and
  `PromotionMetadataTest` now rejects non-`.json`/`.txt` files plus orphan replay-text artifacts
  so temporary refresh leftovers cannot silently re-enter version control.
- Root `./check.sh` stall diagnostics now bound heavyweight per-process captures to a small sample,
  so a badly wedged stage cannot fan out `lsof` or `jcmd` collection across an unbounded
  descendant process tree.
- Supported `jazzer/bin/*` active-fuzz runs now force `--no-daemon` and tear down the launched
  Gradle client tree on interrupt or timeout, so canceling a local fuzz session no longer drops the
  run lock while leaving a live harness JVM and wrapper client chewing CPU in the background.

## [0.33.0] - 2026-04-10

### Changed

- Root build policy is now fully convention-plugin owned. The repository root build script is
  reduced to a thin `gridgrind.root-conventions` application, shared Java module policy lives in
  `GridGrindJavaConventionsPlugin`, and the nested Jazzer build now inherits the same shared
  Spotless and PMD enforcement instead of depending on an old root `subprojects {}` block.
- Nested Jazzer verification now uses dedicated local-only static-analysis profiles: a Jazzer PMD
  ruleset for support and operator code, a fuzz-harness PMD ruleset for `@FuzzTest` entrypoints,
  and a dedicated JaCoCo verification scope for deterministic support-contract classes rather than
  the root product modules' blanket 100% bundle gate.

### Fixed

- Shared build logic no longer recompiles nondeterministically after local edits. The obsolete
  root `buildSrc` directory is gone, build-logic output directories are cleaned without deleting
  the compiler's classpath root, and Kotlin incremental compilation is disabled for
  `gradle/build-logic` so composite-build recompiles no longer lose sibling helper types.
- Root aggregated coverage now resolves reproducibly under configuration cache. The
  `gridgrind.root-conventions` plugin now declares the root repository needed by the
  `jacocoAggregatedReport` task instead of relying on repository state that used to live in the
  old root build script.
- Nested Jazzer support verification is once again fully live under `./gradlew --project-dir
  jazzer check`: the interrupted `XlsxRoundTripVerifierTest` refactor is completed, Jazzer-only
  PMD findings are enforced through the right profile, and regression replay plus support-test
  pulses remain green under the shared conventions.
- Active fuzz scripts once again execute through Jazzer's real command-line JUnit launcher instead
  of a partial in-repo reimplementation. `JazzerHarnessRunner` now requires exactly one
  `@FuzzTest` per harness class, delegates to Jazzer's official `JUnitRunner`, and honors
  `jazzer.max_duration` and `jazzer.max_executions` during local live runs.
- Excel XMLBeans `sqref` handling now goes through one shared normalizer instead of ad hoc raw
  stream usage. Conditional-formatting and data-validation paths both normalize the same way, and
  engine compilation is free of the lingering unchecked-operation notes in the conditional-
  formatting controller and its tests.

## [0.32.2] - 2026-04-10

### Added

- Added [docs/DEVELOPER_GRADLE.md](./docs/DEVELOPER_GRADLE.md), a developer-facing map of the
  Gradle system that explains the shared included build logic, the nested Jazzer composite build,
  the single version-catalog authority, and the periodic review questions contributors should use
  when revisiting the build architecture.

### Changed

- Root and nested Jazzer Gradle builds now share one included build-logic project under
  `gradle/build-logic`, and the nested Jazzer build now imports the root version catalog instead
  of hardcoding overlapping JUnit, Jackson, Apache POI, Log4j, and Jazzer coordinates locally.
- Jazzer harness and run-target metadata now comes from one committed topology file,
  `jazzer/src/main/resources/dev/erst/gridgrind/jazzer/support/jazzer-topology.json`, which is
  consumed by both the runtime support layer and the nested build's task registration.
- `jazzer/build.gradle.kts` is now a thin plugin application rather than a 683-line mixed build
  script, while the nested build still preserves the same public task names and `jazzer/bin/*`
  operator surface.

### Fixed

- `./check.sh` no longer launches stage logging through a racy temporary FIFO that could be
  unlinked before `tee` opened it. Local release verification therefore no longer intermittently
  fails during later stages with `No such file or directory` even when the underlying stage
  command itself succeeds.
- Deleted build-logic helpers can no longer linger as stale hidden `buildSrc` classes in local
  Gradle state. GridGrind now compiles its shared included build logic from clean class output
  directories on each rebuild, and the obsolete root/nested `buildSrc` builds are gone.
- Jazzer support-test pulses and root-project Gradle test pulses now share one scheduled pulse
  foundation, so heartbeat scheduling, thread naming, and whitespace normalization no longer drift
  independently between the two build layers.

## [0.32.1] - 2026-04-09

### Changed

- `./check.sh` now emits explicit per-stage elapsed times plus a total elapsed time in its final
  summary, so long local verification runs show where wall-clock time went without requiring
  external timing wrappers.
- `./check.sh` now derives its Jazzer regression-target progress total from the live regression
  pulse plan instead of a hardcoded harness count, so adding or removing a replay harness cannot
  silently desynchronize the Stage 2 progress summary from the actual nested build.

### Fixed

- Local Docker smoke verification now includes the legal files copied by the production image.
  `.dockerignore` no longer strips those files from the build context, so local image builds match
  the Dockerfile contract instead of failing at copy time.
- Local Docker smoke now asserts GridGrind's help and version output semantically instead of only
  checking that the container process exits successfully, so release-surface contract drift is
  caught before publication.
- Root-project Gradle test pulses now include scheduled in-flight heartbeats for long-running
  tests instead of only reporting progress after completed tests. Local `./check.sh` quality-gate
  monitoring therefore no longer misclassifies healthy long tests as stalled just because a single
  test method runs quietly for longer than the stall threshold.
- Jazzer regression replay now validates promoted-metadata target keys and referenced artifacts
  before replaying them, and mismatch diagnostics now include richer expectation details plus any
  unexpected-failure stack trace. Corrupted or partially moved promoted metadata therefore fails
  fast as a regression-infrastructure defect instead of surfacing as a less actionable replay
  mismatch later.

## [0.32.0] - 2026-04-09

### Changed

- `--help`, the protocol catalog, and the docs now make two important workbook rules explicit:
  Excel sheet names reject reserved characters, and relative `FILE` hyperlinks are checked
  relative to the saved workbook directory during workbook-health analysis.

### Fixed

- Column grouping no longer routes `GROUP_COLUMNS(..., collapsed=false)` through Apache POI's
  collapsed-group expansion path. Overlapping expanded and previously collapsed column groups now
  stay deterministic instead of surfacing the XMLBeans `IndexOutOfBoundsException` that POI can
  trigger while rewriting split column definitions.
- Column outline edits now discard ghost column metadata and canonicalize ambiguous Excel column
  definitions before layout reads and persistence. Repeated no-op ungroup operations therefore no
  longer poison later collapsed groups, and overlapping outline edits keep the same visible state
  across save/reopen cycles instead of drifting when Apache POI leaves stale column definitions in
  memory.
- Sheet-name validation is now consistent across request parsing and engine execution. Invalid
  Excel sheet-name characters and leading or trailing apostrophes are rejected up front with
  structured `READ_REQUEST` failures instead of leaking raw Apache POI sheet-creation errors
  later in `APPLY_OPERATION`.
- Missing relative `FILE` hyperlink findings now explain the workbook directory they were
  resolved against, so the health-check output makes the relative-path anchor obvious instead of
  forcing callers to reverse-engineer it from the final resolved path alone.
- Jazzer now replays the previously crashing engine-command-sequence artifact
  `overlapping_collapsed_then_expanded_group_columns_expected_invalid` as a committed
  expected-invalid regression input, locking the no-crash behavior into the verification suite.

## [0.31.0] - 2026-04-08

### Added

- Jazzer promoted seed `partial_collapsed_column_ungroup_roundtrip_success` now locks in the
  `.xlsx` round-trip case where a collapsed grouped column band is partially ungrouped after save.
- Jazzer support-test progress pulses now include committed class-start and in-flight
  `test-progress` heartbeats, so long-running nested support tests remain observable to the outer
  gate instead of going dark between test completions.
- Added [docs/DEVELOPER_JAVA.md](./docs/DEVELOPER_JAVA.md), which documents GridGrind's shell-level
  Java 26 setup, the `/usr/bin/java` macOS pitfall, the official `jdk.java.net/26` install path,
  and why `./gradlew` is the only supported Gradle entrypoint.
- Added a no-save workbook-health example at
  [examples/workbook-health-request.json](./examples/workbook-health-request.json), showing
  `ANALYZE_WORKBOOK_FINDINGS` as the default lint-style workflow and the correct quoted formula
  syntax for sheet names with spaces.
- Successful protocol responses can now include a `warnings` array. The first shipped warning flags
  formulas that reference same-request sheet names with spaces without single quotes.

### Changed

- Upgraded PMD from 7.22.0 to 7.23.0 and Error Prone from 2.48.0 to 2.49.0 across the shared
  Gradle analyzer toolchain, so all Java subprojects now compile and lint against the newer
  static-analysis baselines.
- Jazzer `.xlsx` round-trip verification now snapshots persisted expectations from the actual
  pre-save workbook state and uses command replay only to bound which cell snapshots are compared
  after reopen. This removes duplicated sheet-layout modeling from the verifier path.
- Local `./check.sh` Stage 1 verification now feeds its stall monitor with semantic
  `[GRADLE-TEST-PULSE]` progress from Gradle `Test` tasks, so long-running root quality-gate runs
  are tracked by actual test execution instead of a stale `> Task :engine:test` banner.
- GridGrind's developer and release docs now treat shell-level Java 26 as a first-class runtime
  requirement and explicitly distinguish that from the temporary JVM 25 bytecode target used only
  by Gradle `buildSrc` logic while Kotlin lacks direct JVM 26 output.
- Gradle `buildSrc` logic now compiles with the Java 26 toolchain while still emitting JVM 25
  bytecode, so local builds no longer require a separate Java 25 installation just to satisfy the
  temporary Kotlin build-logic ceiling.
- CLI help, the README, and the public protocol docs now surface GridGrind's coordinate split much
  more explicitly: `address` and `range` stay in A1 notation, while `*RowIndex` and
  `*ColumnIndex` fields are zero-based and rendered back with Excel-native equivalents in
  validation messages.
- `ANALYZE_WORKBOOK_FINDINGS` is now documented consistently as GridGrind's primary
  workbook-health check, including the no-save `persistence.type=NONE` workflow.

### Fixed

- Jazzer no longer reports a false `.xlsx` round-trip failure for partially ungrouped collapsed
  column groups. GridGrind now accepts Excel's persisted boundary-column collapsed-marker
  semantics for that case, and artifact-backed replay tests load exact committed fixture bytes
  instead of hand-copied inline Base64.
- `./check.sh` no longer false-stalls healthy Stage 1 root verification during quiet but active
  Gradle test execution, and its pulse output now shows live test-progress facts instead of only
  task-start lines.
- `./check.sh` no longer false-stalls healthy Jazzer Stage 2 support-test runs because the
  `JazzerSupportTestPulseListener` is once again source-owned and emits progress during long
  individual support tests instead of relying on stale compiled build state and completion-only
  pulses.
- `./check.sh` now fails fast when the active shell resolves `java` or `javac` to the macOS
  launcher stubs or to anything other than Java 26, so local verification cannot silently run on
  the wrong ambient runtime.
- `EXECUTE_REQUEST` failure context now preserves the parsed request's `sourceType` and
  `persistenceType` instead of dropping them to `null`.
- Row and column bounds errors now report Excel-native equivalents inline, such as
  `firstRowIndex 5 (Excel row 6)` and `firstColumnIndex 5 (Excel column F)`, across structural
  edits, print-title bands, and readback validation records.

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
  fully-qualified class names (e.g., `dev.erst.gridgrind.contract.read.WorkbookReadOperation`) or
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

[Unreleased]: https://github.com/resoltico/GridGrind/compare/v0.68.0...HEAD
[0.68.0]: https://github.com/resoltico/GridGrind/compare/v0.67.0...v0.68.0
[0.67.0]: https://github.com/resoltico/GridGrind/compare/v0.66.0...v0.67.0
[0.66.0]: https://github.com/resoltico/GridGrind/compare/v0.65.0...v0.66.0
[0.65.0]: https://github.com/resoltico/GridGrind/compare/v0.64.0...v0.65.0
[0.64.0]: https://github.com/resoltico/GridGrind/compare/v0.63.0...v0.64.0
[0.63.0]: https://github.com/resoltico/GridGrind/compare/v0.62.0...v0.63.0
[0.62.0]: https://github.com/resoltico/GridGrind/compare/v0.61.0...v0.62.0
[0.61.0]: https://github.com/resoltico/GridGrind/compare/v0.60.0...v0.61.0
[0.60.0]: https://github.com/resoltico/GridGrind/compare/v0.59.0...v0.60.0
[0.59.0]: https://github.com/resoltico/GridGrind/compare/v0.58.0...v0.59.0
[0.58.0]: https://github.com/resoltico/GridGrind/compare/v0.57.0...v0.58.0
[0.57.0]: https://github.com/resoltico/GridGrind/compare/v0.56.0...v0.57.0
[0.56.0]: https://github.com/resoltico/GridGrind/compare/v0.55.0...v0.56.0
[0.55.0]: https://github.com/resoltico/GridGrind/compare/v0.54.0...v0.55.0
[0.54.0]: https://github.com/resoltico/GridGrind/compare/v0.53.0...v0.54.0
[0.53.0]: https://github.com/resoltico/GridGrind/compare/v0.52.0...v0.53.0
[0.52.0]: https://github.com/resoltico/GridGrind/compare/v0.51.0...v0.52.0
[0.51.0]: https://github.com/resoltico/GridGrind/compare/v0.50.0...v0.51.0
[0.50.0]: https://github.com/resoltico/GridGrind/compare/v0.49.0...v0.50.0
[0.49.0]: https://github.com/resoltico/GridGrind/compare/v0.48.0...v0.49.0
[0.48.0]: https://github.com/resoltico/GridGrind/compare/v0.47.0...v0.48.0
[0.47.0]: https://github.com/resoltico/GridGrind/compare/v0.46.0...v0.47.0
[0.46.0]: https://github.com/resoltico/GridGrind/compare/v0.45.0...v0.46.0
[0.45.0]: https://github.com/resoltico/GridGrind/compare/v0.44.0...v0.45.0
[0.44.0]: https://github.com/resoltico/GridGrind/compare/v0.43.0...v0.44.0
[0.43.0]: https://github.com/resoltico/GridGrind/compare/v0.42.0...v0.43.0
[0.42.0]: https://github.com/resoltico/GridGrind/compare/v0.41.0...v0.42.0
[0.41.0]: https://github.com/resoltico/GridGrind/compare/v0.40.0...v0.41.0
[0.40.0]: https://github.com/resoltico/GridGrind/compare/v0.39.0...v0.40.0
[0.39.0]: https://github.com/resoltico/GridGrind/compare/v0.38.0...v0.39.0
[0.38.0]: https://github.com/resoltico/GridGrind/compare/v0.37.0...v0.38.0
[0.37.0]: https://github.com/resoltico/GridGrind/compare/v0.36.0...v0.37.0
[0.36.0]: https://github.com/resoltico/GridGrind/compare/v0.35.0...v0.36.0
[0.35.0]: https://github.com/resoltico/GridGrind/compare/v0.34.0...v0.35.0
[0.34.0]: https://github.com/resoltico/GridGrind/compare/v0.33.0...v0.34.0
[0.33.0]: https://github.com/resoltico/GridGrind/compare/v0.32.2...v0.33.0
[0.32.2]: https://github.com/resoltico/GridGrind/compare/v0.32.1...v0.32.2
[0.32.1]: https://github.com/resoltico/GridGrind/compare/v0.32.0...v0.32.1
[0.32.0]: https://github.com/resoltico/GridGrind/compare/v0.31.0...v0.32.0
[0.31.0]: https://github.com/resoltico/GridGrind/compare/v0.30.0...v0.31.0
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
