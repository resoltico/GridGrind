# Changelog

Notable changes to this project are documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
Earlier release history through `0.68.0` is archived in [docs/CHANGELOG_ARCHIVE.md](docs/CHANGELOG_ARCHIVE.md).

## [Unreleased]

### Changed
- Refreshed build and release dependencies to NullAway `0.13.8`, the Kotlin Gradle plugin `2.4.10`, Jackson Databind `3.2.1`, Bouncy Castle `1.85`, Log4j `2.26.1`, Gradle Shadow `9.5.1`, `actions/setup-java` `5.5.0`, `docker/login-action` `4.4.0`, `docker/setup-buildx-action` `4.2.0`, and `docker/metadata-action` `6.2.0`.

## [0.72.0] - 2026-07-06

### Added
- `--print-recipe-catalog --lookup <id>` now publishes one view-specific recipe detail payload for both built-in examples and task starters, including top-level `intentTags`, the exact runnable `requestProfile`, and task-specific descriptors such as `discoveryProfile`, while the bare catalog remains a compact browse index.
- Scoped protocol-catalog lookups can now publish shared top-level `notes` referenced by stable entry-local `noteRefs`, so reusable rules such as request-owned path handling appear once in machine-readable form instead of being inlined into every matching entry.

### Changed
- CLI execution scratch now uses one private per-run directory instead of request-root `.gridgrind/tmp`: without `--temp-root`, GridGrind creates that scratch directory under the OS temporary-file root; with `--temp-root <path>`, it creates the private per-run scratch directory under the supplied parent path and removes it after the command finishes.
- OOXML write encryption now owns one explicit strong-only request contract: `persistence.security.encryption` is mode-less and AGILE-only, `cipher`/`hash` default to `AES_256` / `SHA_512`, the write allowlist is limited to `AES_256` or `AES_192` plus `SHA_512`, `SHA_384`, or `SHA_256`, and package-security readback continues to report broader factual legacy modes such as `STANDARD` without making them authorable.
- `SET_CONDITIONAL_FORMATTING` now owns its target ranges exactly once: the request target must be `RANGE_BY_RANGE` or `RANGE_BY_RANGES`, the action body carries only the ordered rule definition, and the legacy body-owned `conditionalFormatting.ranges` shape is rejected instead of duplicating the selector contract.
- CLI diagnostics now use one canonical `CliDiagnostic` envelope around the full `problem` core across stderr and stdout-fallback cases, so argument errors, help and discovery failures, request read and validation failures, execution-side stderr diagnostics, and response-write collisions all preserve one problem model plus transport metadata instead of projecting parallel failure schemas.
- Built-in examples and official task starters now derive from one canonical recipe registry, and `--print-recipe-catalog`, `--print-recipe`, and `--print-recipe-keyword-match` now expose example and task views over that same source instead of parallel generators.
- Recipe keyword discovery now ranks against the unified published recipe registry and keeps its browse surfaces compact: list rows stay uniform across example and task views, while detail payloads keep shared concepts such as `intentTags` and `requestProfile` at consistent top-level locations.
- Help, operator, and Docker guidance now treat the bare `--print-protocol-catalog` output as the compact index, route reusable contract rules through scoped `--lookup` `notes`, standardize bind-mounted container commands on `/work`, and teach `--user "$(id -u):$(id -g)"` for ordinary host-owned workspaces.
- Refreshed the container publication workflow action pins to the current GitHub Actions majors used by the release pipeline.

### Fixed
- Encrypted OOXML open/save flows no longer materialize plaintext workbooks under the request tree, execution root, or `--temp-root` parent: decrypted and pre-encryption plaintext temp packages now stay in private OS temp and the CLI no longer leaves residual `.gridgrind/tmp` scratch directories behind after ordinary runs.
- The published Java authoring example and guide no longer teach request-root scratch usage: in-process examples now bind one caller-owned private temp root explicitly, recommend placing it outside the request workspace, and show that callers should clean it up after execution.
- `GET_PACKAGE_SECURITY` now validates each OOXML signature part before reading its signer certificate, so signed workbook readback once again includes factual `signer { subject, issuer, serialNumberHex }` identity on valid signatures instead of dropping signer metadata while still reporting `state=VALID`.
- OOXML-encrypted workbook persistence no longer relies on Apache POI's weak preset constructor: GridGrind now builds the AGILE write envelope explicitly, default encrypted fixtures and parity assets read back as `AES_256` / `SHA_512`, and implicit encrypted-source preservation now refuses to carry forward legacy or unsupported source envelopes without an explicit supported write choice.
- Cell error literals now split cleanly by direction: authored inputs accept only the seven valid stored OOXML error tokens, while readback and formula evaluation report a nine-token GridGrind-owned vocabulary that also covers `#CIRCULAR_REF!` and `#FUNCTION_NOT_IMPLEMENTED!`, preventing Apache POI sentinels from leaking either to the wire or into saved workbook XML.
- CLI recovery suggestions no longer fabricate the unrelated `--print-protocol-catalog --search "sheet layout"` command, response-path fallback diagnostics now stay transport-only instead of duplicating facts already owned by `problem` and `problem.context`, and recipe no-match results now return the published intent-tag vocabulary instead of an empty recovery surface.
- SUMMARY execution journals now retain their published compact per-step summaries instead of dropping `journal.steps` entirely, and repeated SUMMARY responses serialize byte-identically because step and phase timing telemetry stays stripped while resolved targets and outcomes remain intact.

## [0.71.0] - 2026-07-02

### Added
- Added machine-readable projection metadata to the protocol catalog: the compact index now publishes `fieldMetadataKeys`, `CellReadProjection.facets` publishes `enumValueDocs`, and facet-gated cell-report fields publish `projectedByFacets`, so agents can derive request-to-field mappings such as `VALUE -> textValue|numberValue`, `FORMAT -> displayValue`, `RICH_TEXT_RUNS -> runs`, and `TEMPORAL -> temporal` from the catalog itself.
- Added a richer public cell-readback contract across `GET_CELLS`, `GET_WINDOW`, `GET_SHEET_SCHEMA`, and assertion observations: the readback vocabulary now uses `TEXT`, numeric date-like cells can project `temporal { isDate, kind, isoValue }`, and rich-text runs are exposed as an opt-in projection facet instead of a standalone readback discriminator.

### Changed
- Documented the shipped JSON examples and generated package-security workbook as CLI-owned derived artifacts regenerated from the built-in example registry via `./scripts/sync-generated-examples.sh`, so repository fixtures, packaged discovery output, and verification all point at one source of truth.
- JSON-native CLI payloads are now compact by default across request templates, discovery output, doctor reports, execution responses, and structured identity surfaces; `--pretty` is the explicit opt-in for indented JSON, while `--format text|structured` now applies only to prose-oriented help, version, and license surfaces.
- Protocol discovery now revolves around the compact `--print-protocol-catalog` index plus summary-first `--search` and scoped `--lookup`; the older monolithic `--full` catalog dump is no longer part of the CLI grammar.
- Request envelopes now carry the sparse wire contract all the way through nested execution policy: `execution.mode`, `execution.journal`, and `execution.calculation` default independently when omitted, emitted request templates stay minimal, and the public docs/examples no longer reintroduce boilerplate execution-policy blocks.
- Execution responses now publish one canonical top-level persistence outcome on every success and failure, `SAVE_AS` requires explicit `ifExists=REJECT|REPLACE`, `REPLACE` enables create-or-replace writes, and shipped save-producing examples and task starters now use rerunnable `REPLACE` output.
- Cell-returning reads now default to smaller, more explicit payloads: `GET_WINDOW` is sparse unless `includeBlanks=true`, `GET_CELLS`, `GET_WINDOW`, and `GET_SHEET_SCHEMA` share `projection.facets`, facet-gated fields such as `style`, `displayValue`, `hyperlink`, `comment`, `formula`, `runs`, and `temporal` appear only when requested, and the public examples/templates now default row and range payloads to the canonical `TYPED`/`cells` wrappers.
- Container and operator guidance now standardizes on the image's prepared `/work` directory, so Docker examples, help surfaces, and verification all teach one mounted-working-directory contract instead of ad hoc `-w` overrides.
- Tightened the release control plane so archiving older root changelog history now pairs with deleting any no-longer-widening `release-ledger` exception, keeping the live release ledger within the default governance budget instead of carrying stale reviewed carve-outs forward.

### Fixed
- Rebuilt request intake, doctor batching, and CLI read-failure reporting around typed request-problem descriptors plus a structural Jackson intake detector that derives missing-member failures from the effective creator/discriminator contract instead of downstream prose re-parsing, covering missing required fields, missing type discriminators, explicit `null` placeholders, unknown fields, and nested step payload shape failures.
- Tightened request diagnostics so `context.jsonPath`, public messages, and resolutions stay pinned to the exact offending request value across nested step payloads, duplicate `stepId` detection, non-`.xlsx` workbook path validation, unknown-field failures, and execute-surface step-envelope and target-shape reconstruction without duplicated path segments.
- Aligned the request DTO, doctor, and fuzzing stack with the documented omission defaults for `execution`, `formulaEnvironment`, and OOXML persistence-security blocks, including per-axis execution normalization and serialization so sparse wire payloads now parse and re-emit exactly as documented instead of reviving implicit boilerplate defaults.
- Persist-workbook failures no longer serialize a second `problem.context.persistence` object alongside the canonical top-level response `persistence` outcome, and the failure context now points directly at `sourceWorkbookPath` or `persistencePath`.
- Locked the schema/readback vocabulary end to end: schema `observedTypes` and `dominantType` now reject stale non-canonical tokens such as `STRING`, Jazzer and round-trip carriers now speak the same `TEXT`-first public contract, and compact-payload regression tests now pin the default one-cell and sparse-window response sizes.
- Corrected the public protocol catalog for projection-aware cell reads so defaulted `projection` and `includeBlanks` fields are no longer over-declared as required, facet-gated readback fields no longer masquerade as always present, and lookup payloads now match the runtime omission and defaulting rules.
- Rewrote the `GET_WINDOW` and `GET_SHEET_SCHEMA` limit rationale to match the shipped sparse-window contract, and `GET_CELLS` now shares the same deterministic `250,000`-cell cap so all cell-returning read surfaces reject oversized requests before workbook IO instead of leaving exact-address reads effectively unbounded.

## [0.70.0] - 2026-06-30

### Added
- Added `UNSUPPORTED_FORMULA_CONSTRUCT` as a public problem code for authored formulas that parse successfully but rely on unsupported constructs such as `LAMBDA` and `LET`.

### Changed
- Request envelopes can now omit the default `execution` and `formulaEnvironment` blocks; emitted request templates, built-in examples, task starters, and protocol docs now use the minimal envelope while still defaulting omitted blocks to `FULL_XSSF` / `SUMMARY` / `DO_NOT_CALCULATE` and an empty evaluator environment.
- The packaged shadow distribution is now the only install/archive path: the legacy thin `installDist` / `distZip` / `distTar` tasks no longer materialize a second launcher tree, `installShadowDist` cleans leftover thin-script artifacts, and the README, quick-start guidance, and CLI distribution verification now all describe and test one canonical `gridgrind` launcher contract instead of split launcher paths.
- Refreshed the shared maintenance baseline to Gradle `9.6.1`, JUnit `6.1.1`, NullAway `0.13.7`, Spotless `8.7.0`, `actions/checkout` `7.0.0`, `actions/setup-java` `5.3.0`, and `gradle/actions` `6.2.0`, and taught Dependabot to stop opening duplicate root-wrapper PRs from the nested `/jazzer` build.
- The public error catalog now gives each problem code cause-specific resolution text instead of falling back to one generic recovery message.
- Refreshed the developer and operator docs to the live Gradle `9.6.1`, Jackson `3.2.0`, JUnit `6.1.1`, Log4j `2.26.0`, and doctor-request batching contract so first-contact guidance and contributor references match current HEAD.

### Fixed
- Help, `--help-protocol`, and `--response` prose now match the real stdout/stderr contract: transport and argument failures emit structured JSON on stderr, executed request failures remain primary stdout payloads, and `--response` write-fallback notices now describe the structured failure report that is actually written.
- Explicit `null` placeholders now produce the dedicated message `Field '<x>' must be omitted when absent; explicit null is not accepted.`, and `--doctor-request` now reports both top-level omission-legal `execution` and `formulaEnvironment` null violations in one report instead of stopping after the first one.
- Request doctor now batches independently provable blocking problems across request-default preflight, multiple malformed step payloads, and semantic validation while normal execution still stops at the first blocking failure.
- Nested request-shape diagnostics now keep `location.jsonPath` pinned to the offending value itself, including step envelope, selector target, and payload-shape errors, instead of collapsing failures onto a broader parent object.
- Authored unsupported formulas such as `LAMBDA` and `LET` now classify as `UNSUPPORTED_FORMULA_CONSTRUCT` instead of `INVALID_FORMULA`, and assertion and save-as I/O failures now return recovery guidance that matches the actual cause.

## [0.69.0] - 2026-06-14

### Changed
- Extended the structural-governance stack beyond handwritten Java: root `check` now runs `verifyControlPlaneShape`, which scans repository-owned shell gates, Kotlin build-logic, the release protocol, and the public changelog ledger against expiring reviewed budgets from `gradle/control-plane-shape-policy.tsv` so the repo-governing control plane cannot drift outside the same no-god-file ratchet model as product code.
- Rotated the public release ledger into two owned surfaces: root `CHANGELOG.md` now carries the unreleased stream plus recent releases only, while older release history lives in `docs/CHANGELOG_ARCHIVE.md` so release-time governance tightens around the live operator-facing ledger instead of broadening the control-plane budget whenever the cumulative archive grows.
- Tightened the forbidden tagged-union and god-record enforcement so the build now inspects named nested record/class variants as well as direct top-level declarations, closing the blind spot around the sealed-interface-with-nested-record style GridGrind uses for most domain surfaces.
- Split the packaged CLI discovery and help surface across narrower role-owned seams: identity commands, example/task discovery, protocol-catalog output, trailing-argument validation, and help-section rendering now live on dedicated helpers instead of one broad catalog-command implementation.
- Hardened the release protocol around worktree-driven publication: the documented flow now tells operators to archive older changelog history before widening the live release ledger, and to delete temporary bootstrap branches and bootstrap manifests once the primary checkout has been reconciled back to the published `main` state.

### Fixed
- Finished the OOXML encryption, custom-XML, and grouped table-report hard-break migration across the remaining Jazzer and parity verification surfaces, including invariant checks, promoted test fixtures, and the remaining support tests, so the regression and replay stack now validates the live sealed DTO model instead of stale flat constructors.
- Aligned invalid-request-shape reporting across the CLI doctor, runtime problem surface, and JSON codec layer: missing required root fields and explicit `null` placeholders now classify as `INVALID_REQUEST_SHAPE`, persist-workbook collision reporting no longer dereferences absent save-as paths while building public diagnostics, and the promoted Jazzer protocol-request replay metadata now refreshes to the same decode-outcome truth instead of preserving stale `INVALID_REQUEST` expectations.
- Hardened deterministic OOXML artifact persistence: the package-copy helper now validates its target path explicitly, and deterministic ZIP-package rewrites now clean their temporary output artifact before rethrowing any rewrite failure.
- Repaired the Jazzer workbook IO seam after the explicit write-disposition hard break: replay and round-trip support now save through `WorkbookArtifactWriteDisposition.REPLACE_EXISTING` instead of calling the removed two-argument persistence API.
- Routed default CLI failure reports to stderr instead of stdout, so bare invocation and other transport-level argument failures no longer masquerade as first-class primary payloads on the success channel.
- Standardized the packaged discovery contract around `requestFileName` plus `requiredWorkspacePaths`, and realigned the release verifier, operator guidance, and public docs to that explicit example/task portability surface instead of carrying forward stale `suggestedRequestPath` and `requiredPaths` terminology.
- Made the Docker runtime cache layout arbitrary-user-safe: the image now points `HOME` and `XDG_CACHE_HOME` at writable tmp-backed directories so signature-line and other font-backed authoring flows stay silent under `docker run --user <uid>:<gid>` instead of leaking Fontconfig cache warnings on stderr.

[Unreleased]: https://github.com/resoltico/GridGrind/compare/v0.72.0...HEAD
[0.72.0]: https://github.com/resoltico/GridGrind/compare/v0.71.0...v0.72.0
[0.71.0]: https://github.com/resoltico/GridGrind/compare/v0.70.0...v0.71.0
[0.70.0]: https://github.com/resoltico/GridGrind/compare/v0.69.0...v0.70.0
[0.69.0]: https://github.com/resoltico/GridGrind/compare/v0.68.0...v0.69.0
