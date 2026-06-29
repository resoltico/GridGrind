# Changelog

Notable changes to this project are documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
Earlier release history through `0.64.0` is archived in [docs/CHANGELOG_ARCHIVE.md](docs/CHANGELOG_ARCHIVE.md).

## [Unreleased]

### Fixed

- Help now documents the real stderr/stdout split, executed failure responses stay on stdout, and
  explicit `null` request placeholders surface a dedicated “omit the field instead” shape error.

## [0.69.0] - 2026-06-14

### Changed

- Extended the structural-governance stack beyond handwritten Java: root `check` now runs
  `verifyControlPlaneShape`, which scans repository-owned shell gates, Kotlin build-logic, the release
  protocol, and the public changelog ledger against expiring reviewed budgets from
  `gradle/control-plane-shape-policy.tsv` so the repo-governing control plane cannot drift outside the same
  no-god-file ratchet model as product code.
- Rotated the public release ledger into two owned surfaces: root `CHANGELOG.md` now carries the
  unreleased stream plus recent releases only, while older release history lives in
  `docs/CHANGELOG_ARCHIVE.md` so release-time governance tightens around the live operator-facing
  ledger instead of broadening the control-plane budget whenever the cumulative archive grows.
- Tightened the forbidden tagged-union and god-record enforcement so the build now inspects named nested
  record/class variants as well as direct top-level declarations, closing the blind spot around the
  sealed-interface-with-nested-record style GridGrind uses for most domain surfaces.
- Split the packaged CLI discovery and help surface across narrower role-owned seams: identity commands,
  example/task discovery, protocol-catalog output, trailing-argument validation, and help-section
  rendering now live on dedicated helpers instead of one broad catalog-command implementation.
- Hardened the release protocol around worktree-driven publication: the documented flow now tells
  operators to archive older changelog history before widening the live release ledger, and to
  delete temporary bootstrap branches and bootstrap manifests once the primary checkout has been
  reconciled back to the published `main` state.

### Fixed

- Finished the OOXML encryption, custom-XML, and grouped table-report hard-break migration across the
  remaining Jazzer and parity verification surfaces, including invariant checks, promoted test fixtures, and
  Stage 3 support tests, so the regression and replay stack now validates the live sealed DTO model instead
  of stale flat constructors.
- Aligned invalid-request-shape reporting across the CLI doctor, runtime problem surface, and JSON codec
  layer: missing required root fields and explicit `null` placeholders now classify as `INVALID_REQUEST_SHAPE`,
  persist-workbook collision reporting no longer dereferences absent save-as paths while building public diagnostics,
  and the promoted Jazzer protocol-request replay metadata now refreshes to the same decode-outcome truth instead of
  preserving stale `INVALID_REQUEST` expectations.
- Hardened deterministic OOXML artifact persistence: the package-copy helper now validates its target path
  explicitly, and deterministic ZIP-package rewrites now clean their temporary output artifact before rethrowing
  any rewrite failure.
- Repaired the Jazzer workbook IO seam after the explicit write-disposition hard break: replay and
  round-trip support now save through `WorkbookArtifactWriteDisposition.REPLACE_EXISTING` instead
  of calling the removed two-argument persistence API.
- Routed default CLI failure reports to stderr instead of stdout, so bare invocation and other
  transport-level argument failures no longer masquerade as first-class primary payloads on the
  success channel.
- Standardized the packaged discovery contract around `requestFileName` plus
  `requiredWorkspacePaths`, and realigned the release verifier, operator guidance, and public docs
  to that explicit example/task portability surface instead of carrying forward stale
  `suggestedRequestPath` and `requiredPaths` terminology.
- Made the Docker runtime cache layout arbitrary-user-safe: the image now points `HOME` and
  `XDG_CACHE_HOME` at writable tmp-backed directories so signature-line and other font-backed
  authoring flows stay silent under `docker run --user <uid>:<gid>` instead of leaking Fontconfig
  cache warnings on stderr.

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

[Unreleased]: https://github.com/resoltico/GridGrind/compare/v0.69.0...HEAD
[0.69.0]: https://github.com/resoltico/GridGrind/compare/v0.68.0...v0.69.0
[0.68.0]: https://github.com/resoltico/GridGrind/compare/v0.67.0...v0.68.0
[0.67.0]: https://github.com/resoltico/GridGrind/compare/v0.66.0...v0.67.0
[0.66.0]: https://github.com/resoltico/GridGrind/compare/v0.65.0...v0.66.0
[0.65.0]: https://github.com/resoltico/GridGrind/compare/v0.64.0...v0.65.0
