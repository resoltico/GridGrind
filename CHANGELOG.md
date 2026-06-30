# Changelog

Notable changes to this project are documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
Earlier release history through `0.65.0` is archived in [docs/CHANGELOG_ARCHIVE.md](docs/CHANGELOG_ARCHIVE.md).

## [Unreleased]

### Added
- Added `UNSUPPORTED_FORMULA_CONSTRUCT` as a public problem code for authored formulas that parse successfully but rely on unsupported constructs such as `LAMBDA` and `LET`.

### Changed
- Request envelopes can now omit the default `execution` and `formulaEnvironment` blocks; emitted request templates, built-in examples, task starters, and protocol docs now use the minimal envelope while still defaulting omitted blocks to `FULL_XSSF` / `SUMMARY` / `DO_NOT_CALCULATE` and an empty evaluator environment.
- The packaged shadow distribution is now the sole owner of the canonical `gridgrind` launcher name, and the README, quick-start guidance, example guide, and CLI distribution verification now all describe and test that single launcher contract instead of splitting guidance across competing launcher paths.
- Refreshed the shared maintenance baseline to Gradle `9.6.1`, JUnit `6.1.1`, NullAway `0.13.7`, Spotless `8.7.0`, `actions/checkout` `7.0.0`, `actions/setup-java` `5.3.0`, and `gradle/actions` `6.2.0`, and taught Dependabot to stop opening duplicate root-wrapper PRs from the nested `/jazzer` build.
- The public error catalog now gives each problem code cause-specific resolution text instead of falling back to one generic recovery message.

### Fixed
- Help, `--help-protocol`, and `--response` prose now match the real stdout/stderr contract: transport and argument failures emit structured JSON on stderr, executed request failures remain primary stdout payloads, and `--response` write-fallback notices now describe the structured failure report that is actually written.
- Explicit `null` placeholders now produce the dedicated message `Field '<x>' must be omitted when absent; explicit null is not accepted.`, and `--doctor-request` now reports both top-level omission-legal `execution` and `formulaEnvironment` null violations in one report instead of stopping after the first one.
- Request doctor now returns every independently provable semantic validation problem in one machine-readable response while normal execution still stops at the first blocking semantic failure.
- Nested request-shape diagnostics now keep `location.jsonPath` pinned to the offending value itself, including step envelope, selector target, and payload-shape errors, instead of collapsing failures onto a broader parent object.
- Authored unsupported formulas such as `LAMBDA` and `LET` now classify as `UNSUPPORTED_FORMULA_CONSTRUCT` instead of `INVALID_FORMULA`, and assertion and save-as I/O failures now return recovery guidance that matches the actual cause.

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

[Unreleased]: https://github.com/resoltico/GridGrind/compare/v0.69.0...HEAD
[0.69.0]: https://github.com/resoltico/GridGrind/compare/v0.68.0...v0.69.0
[0.68.0]: https://github.com/resoltico/GridGrind/compare/v0.67.0...v0.68.0
[0.67.0]: https://github.com/resoltico/GridGrind/compare/v0.66.0...v0.67.0
[0.66.0]: https://github.com/resoltico/GridGrind/compare/v0.65.0...v0.66.0
