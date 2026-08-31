# Changelog

Notable changes to this project are documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
Earlier release history through `0.68.0` is archived in [docs/CHANGELOG_ARCHIVE.md](docs/CHANGELOG_ARCHIVE.md).

## [Unreleased]

### Added
- Added `--materialize-recipe --lookup <id> --workspace <new-directory>` to publish an asset-backed recipe into one new workspace atomically, including its request file and every declared asset.
- Added a dedicated `check_mutation.sh` entrypoint and scheduled PIT `1.30.0` workflow for the reviewed contract, engine, CLI, and architecture-rule scopes; it validates configured class and test patterns against compiled bytecode, requires all four reports, and accepts only killed mutations.
- Added ArchUnit `1.5.0` architecture verification to the normal repository gate, with an independent ArchUnit/Jupiter task pair that imports only the five product modules and enforces the fail-closed rule inventory for dependency direction, public-surface isolation, closed model variants, formula ownership, and POI access.
- Added fail-closed runtime legal-inventory verification to both CLI checking and fat-JAR packaging so any resolved dependency or version drift requires an explicit artifact-license and NOTICE review before it can ship.

### Changed
- Updated NullAway to `0.14.1`.
- `--print-recipe` now emits only self-contained recipes; asset-backed recipes direct callers to `--materialize-recipe`, which requires a new workspace and does not accept `--response`.
- Formula autosizing now uses existing cached results or formula text without evaluating formulas or changing cached workbook values.
- Invalid JSON diagnostics now include one-based line and column coordinates alongside UTF-8 byte offsets, and unknown type diagnostics list the valid values for their request union.
- Generated named-range templates now use the matching `RevenueRange` placeholder.
- Refactored the internal protocol-catalog lookup model into a closed record family with centralized lookup presentation and search behavior, preserving the published lookup payload while making the model boundary explicit.
- Expanded Jazzer generation and deterministic replay with explicit-list boundary cases, delimiter-rejection cases, oversized row-work cases, and CR-only syntax inputs.

### Fixed
- Corrected the distributed legal inventory and packaging: the Eclipse Distribution License v1.0 terms for Jakarta Activation and Jakarta XML Binding are now reproduced with their component-specific notices and identified by their `BSD-3-Clause` SPDX identifier; required upstream NOTICE attributions and Jackson Core's shaded FastDoubleParser, Schubfach, MIT, Boost 1.0, and BSD 2-Clause terms are retained; the source-only Gradle wrapper and the Apache POI Custom XML workbook embedded as a packaged recipe asset are attributed with exact provenance and distribution scope; Gradle-generated launcher archives expose legal files directly; the standalone release JAR embeds them and identifies its aggregate license posture as multiple rather than MIT-only; patent prose no longer implies an unperformed clearance or separate non-assertion pledge; and the container no longer publishes an incomplete aggregate OCI license expression.
- Corrected the contributor foundations table so its Jackson Databind, JUnit Jupiter, Log4j Core, and Kotlin versions match the canonical Gradle version catalog.
- Corrected contributor references that still described the published runtime image as Alpine-based and named superseded Gradle and JUnit versions.
- Rejected XML 1.0-invalid authored cell and rich-text characters before workbook mutation instead of silently replacing them during OOXML persistence.
- Rejected reversed ranges, range/grid dimension mismatches, and oversized cell- or row-materializing plans before Apache POI authoring.
- Rejected conditional-formatting rules that neither apply a style nor act as a `stopIfTrue` barrier, and made CLI request rejections use the documented exit status consistently.
- Corrected the Docker build context so packaged asset-backed recipes, including custom-XML workflows, retain their required workspace assets in the published image.
- Corrected explicit data-validation lists so delimiter-bearing values are rejected rather than split, and the 255-character limit applies to the complete stored formula.
- Corrected CR-only JSON diagnostics to report the right line and column.
- Corrected the generated `SET_RANGE` catalog template so its target and typed grid dimensions agree.

## [0.74.0] - 2026-08-28

### Added
- Added machine-readable protocol-catalog scalar constraints and operation preconditions, including exact formats, lengths, numeric bounds, integral values, required nonblank text, and `.xlsx` path suffixes where the public contract enforces them.
- Added caller-actionable diagnostics for circular formula evaluation, source and output paths that name directories, create-new output collisions, and a leading UTF-8 byte-order mark; diagnostics retain the relevant authored location when one exists.
- Added typed response-transport reasons for an existing, directory, unwritable, or late-failing `--response` destination so automation can distinguish an undelivered result from a result written to stdout.

### Changed
- Updated the Gradle Wrapper to `9.7.1`; the aligned JUnit BOM, Jupiter, and Platform Launcher to `6.1.3`; Jackson Databind to `3.2.2`; the Error Prone Gradle plugin to `5.1.1`; and NullAway to `0.14.0`.
- Replaced nullable `CellStyleReport.border` side facts with explicit `NONE`, `DEFAULT_COLOR`, and `COLORED` variants, allowing exact style readback to feed `EXPECT_CELL_STYLE` without null padding.
- Replaced flat assertion outcomes with tagged `PASSED` and `FAILED` variants; each failed entry in `assertions[]` now carries its complete target, authored assertion, and observed evidence.
- Replaced conditional-formatting threshold payloads with explicit `MIN`, `MAX`, `NUMBER`, `PERCENT`, `PERCENTILE`, and `FORMULA` variants, separating authorable threshold inputs from factual readback-only states.
- Replaced generic workbook-qualified cell addresses in formula evaluation targets with the dedicated `FormulaCellTarget` model and removed the unreachable `CELL_BY_QUALIFIED_ADDRESSES` selector from the request contract.
- Aligned named-range operation targets, generated step templates, doctoring, execution, and protocol-catalog metadata on the same scope-specific selector contract; table-row selectors remain supported nested inputs.
- Changed normal `FORMULA` authoring to validate after preceding mutations are present and to retain the authoring step when a later operation surfaces a formula failure; `RAW_FORMULA` remains intentionally opaque.
- Changed `--response` execution to reserve a new no-follow destination after request validation and before input binding or workbook work, preserving the requested create-new transport contract without replacing an existing response file.
- Improved protocol-catalog discovery so exact identifiers rank first, partial multi-term searches remain useful, and `FILE` hyperlink documentation explicitly states that relative targets are resolved from the saved workbook's directory.

### Fixed
- Corrected all data-validation comparison operators so their persisted OOXML and reopened workbook semantics match the authored Excel rule.
- Corrected defined-name and table-name validation to accept supported Unicode identifiers, while preserving factual readback of names already stored in existing workbooks.
- Corrected named-formula health analysis to inspect workbook-aware formula references, handling quoted sheet names, escaped apostrophes, multiple references, and unparseable formulas without guessing sheet names.
- Corrected calculation handling for circular formulas: strict evaluation reports `CIRCULAR_FORMULA_REFERENCE`, while lenient calculation reports the unevaluable formula without treating ordinary evaluated Excel errors as unevaluable.
- Corrected factual formula inspections so `GET_FORMULA_SURFACE` and `GET_CELLS` with only the `FORMULA` facet can read `RAW_FORMULA` cells without requesting evaluation.
- Corrected doctoring and execution preflight to preserve caller-correctable request-size and request-path codes, detect ordered formula-and-column-edit conflicts, and reject invalid output leaves before workbook mutation.
- Corrected response-file handling so an existing requested response path stops execution before workbook side effects and cannot leave a stale payload reported as a successful execution result.
- Corrected request diagnostics to retain exact nested collection and tagged-union JSON paths and byte offsets instead of collapsing failures onto a broader parent or an unavailable location.
- Corrected source and output path classification, create-new collision handling, and encrypted-workbook diagnostics so public responses use authored paths rather than private materialization paths.
- Corrected `SAVE_AS` to create missing contained destination parents while retaining fail-closed no-follow path checks.
- Corrected leading UTF-8 byte-order mark handling: exactly one mark at byte zero is ignored with `UTF8_BOM_IGNORED`, and later marks remain invalid JSON.
- Corrected the shared drawing and signature PNG asset so generated workbooks open in LibreOffice without a media CRC warning.

## [0.73.0] - 2026-08-21

### Added
- Added the public conformance record for deterministic responses, fail-closed request-path capability, and the explicitly environment-sensitive signature-preview boundary.
- Added compact JSONL live progress on stderr for `execution.journal.level=VERBOSE`. Each event carries a timestamp, lifecycle category, status, and optional authored step identity; `FAILED` events additionally carry their problem code. Events never carry free-form detail or declared-secret values, and `--pretty` affects only the primary payload.
- Added `execution.assertionMode=COLLECT` for terminal verification phases: every assertion is evaluated and reported before the first canonical assertion failure is returned; static validation rejects any mutation after the first collected assertion.
- Added `WORKBOOK_NOT_OPENABLE` for corrupted, truncated, non-zip, or non-workbook OOXML source packages and `ENCRYPTION_SOURCE_NOT_PRESERVABLE` for encrypted source envelopes outside the AGILE write contract, separating explicit request-policy failures from cryptographic failures.
- Added `RAW_FORMULA` for opaque OOXML formula-body authoring when newer Excel syntax cannot be parsed by POI. Formula character data is XML-safe across full and streaming writes, while forbidden XML controls and invalid framing are rejected as `INVALID_FORMULA_TEXT`.
- Structured request warnings now carry a typed `location`, distinguishing step-owned warnings from request-path warnings; contained absolute request paths emit `NON_PORTABLE_ABSOLUTE_PATH` with their exact path role.
- Request doctoring now reports every independently observable request-intake defect in one response, including duplicate keys, unknown fields, omitted required fields, explicit nulls, malformed scalar values, missing or unknown type discriminators, and constructor-level field validation failures; valid sibling fragments remain available for safe preflight rather than being discarded after the first defect.
- Protocol catalog field descriptors now publish `secret: true` for authored credential-bearing fields, allowing consumers to identify values that diagnostics and telemetry must never reproduce.
- Added the plural `CommandError` result for command failures before workbook execution, with `status=REJECTED` and a nonempty canonical `problems[]` collection. Request locations retain exact UTF-8 offsets and duplicate-property identity where the source contains an offending token.
- Added stable `GridGrindWarningCode` values to every structured request warning so warning consumers no longer need to infer meaning from prose alone.

### Changed
- Existing-workbook persistence now requires an explicit total OOXML security policy: encryption is `NONE`, `ENCRYPT`, or `PRESERVE_SOURCE`; signature is `NONE` or `SIGN`; a writing existing source must declare both axes.
- The Docker image now uses the glibc-based Zulu 26 runtime so mounted workspaces retain the required secure no-follow directory handles for request-owned output paths; the container workflow remains fail-closed when that capability is unavailable.
- Formula text now uses one exact OOXML `<f>`-body convention: `FORMULA` and `RAW_FORMULA` values must not begin with `=`. Lenient `EVALUATE_ALL` and `EVALUATE_TARGETS` keep unevaluable formulas unchanged, report `PARTIAL`, and emit `FORMULA_NOT_EVALUATED`; `REQUIRE_EVALUATION` is the explicit strict strategy, while `DEFERRED_CALCULATION` reports capabilities without attempting server-side evaluation.
- Static request validity now has one contract layer shared by request analysis, doctoring, normal execution, and protocol-catalog target discovery. A known selector shape now binds independently of the operation that received it, and an incompatible pair is reported at `steps[i].target.type` as `INVALID_REQUEST` instead of being mislabeled as a malformed request shape. Execution-mode compatibility, calculation ordering, and persistence compatibility use that same static contract rather than separate runtime rules.
- The request and response contract now uses protocol version `V2` exclusively; `V1` requests are rejected. Request-side tagged unions use the uniform `type` discriminator, including colors, fills, drawing shapes, and named-range targets, while unrelated report-domain `kind` fields retain their established meanings.
- Protocol-catalog field requirements now derive from each request record's effective JSON creator contract and selected-union discriminator, including explicit defaultable fields and JSON property names, instead of catalog-local required or optional field lists.
- Optional boolean request fields now declare their effective wire default in the request model and protocol catalog as `defaultBoolean`, so omission resolves consistently while explicit `null` remains invalid.
- Sensitive request diagnostics now use the contract's exact secret-owner JSON paths rather than global text replacement, so a problem at a credential path is safely generic without corrupting unrelated workbook data or diagnostic text that happens to contain a short secret; workbook, revisions, and OOXML encryption passwords are declared as secret fields, and last-resort failures use the canonical internal-error title instead of arbitrary throwable text.
- Refreshed build and release dependencies to NullAway `0.13.8`, the Kotlin Gradle plugin `2.4.10`, JUnit `6.1.2`, Jackson Databind `3.2.1`, Bouncy Castle provider `1.85.2` (with current PKIX/util artifacts at `1.85`), Log4j `2.26.1`, JSpecify `1.0.1`, Google Java Format `1.36.1`, PMD `7.26.0`, Gradle Shadow `9.6.1`, Spotless `8.10.0`, `actions/checkout` `7.0.1`, `actions/setup-java` `5.7.0`, Gradle Actions `6.3.0`, `docker/login-action` `4.6.0`, `docker/setup-buildx-action` `4.3.0`, and `docker/metadata-action` `6.2.0`.
- Style authoring now names partial updates honestly: the catalog and request model use `CellStylePatchInput` for `APPLY_STYLE.style`, while the separate complete `CellStyleReport` remains the factual readback shape. Cell-style and conditional-formatting patches now share `BorderSideInput` and `ColorInput`, so RGB, theme, indexed, and tint color references have one consistent authoring vocabulary.
- Command failures now use `CommandError` directly instead of a parallel CLI diagnostic wrapper. `REJECTED` describes only that workbook execution never began; the canonical problem's category continues to describe the fault domain.
- Execution results now use `WorkbookResult` exclusively. The top-level result owns optional `planId`, common journal, persistence, warning, assertion, and inspection fields on both outcomes, and only `FAILED` carries the singular canonical `problem`.
- Normal CLI output now has one primary channel: stdout without `--response`, or the requested response file with it. The CLI no longer mirrors an equivalent failure payload on stderr.
- `VERBOSE` now streams its fine-grained lifecycle progress as compact JSONL to stderr instead of duplicating events in `WorkbookResult.journal`; response-file fallback remains one structured transport notice on that same channel.
- Recipe-catalog entries now publish the structured `advisory` field with `requiredWorkspacePaths[]` instead of `workspaceMode`, making self-contained and asset-required preparation directly discoverable without prose parsing.

### Fixed
- Docker builds now pin the Zulu builder and runtime as verified multi-architecture manifest lists, and the smoke gate rejects a base image without both `linux/amd64` and `linux/arm64/v8` variants before building. This prevents architecture-specific `/bin/sh` execution failures on supported hosts.
- Existing-workbook writes no longer inherit encryption or signatures by omission. `PRESERVE_SOURCE` now rejects plaintext or write-incompatible encrypted sources during non-mutating preflight, signing material is fully validated before workbook mutation, `signature: NONE` physically removes source signatures, and `SIGN` replaces them with a fresh signature.
- Request intake now rejects non-finite or precision-losing IEEE numeric literals as `NUMBER_NOT_REPRESENTABLE` at their exact JSON path and UTF-8 token offset, while accepting ordinary decimal forms such as `0.1`, `1e0`, and `100.00`.
- Request-owned workbook, persistence, formula-environment, signing-material, and file-backed input paths now share one fail-closed boundary: absolute and relative escapes are rejected, all reads are materialized through no-follow descriptor bindings before mutation, and output commits recheck the bound directory topology before writing. Symlinks are never followed, and the documented residual concurrent-topology window is explicit rather than silently overstated.
- Request doctoring now batches independent source-backed-input and existing-workbook preflight failures whenever a complete request is available, including alongside static contract findings; execution performs the same preflight only after static validation passes, and any preflight failure completes zero steps and persists no workbook. Source-backed input failures retain their exact authored `context.json.jsonPath`, and execution selects the same ordered primary preflight problem that doctoring reports.
- Request diagnostics now distinguish `READ_REQUEST` structural intake from `BIND_REQUEST` creator-contract failures, preserving phase-first deterministic ordering even when a constructor failure appears earlier in the JSON payload.
- Asset-backed recipe printing now keeps stderr free of ad hoc portability prose; callers discover the required workspace assets through the catalog's structured `advisory` and `requiredWorkspacePaths` fields, and response-file fallback emits only its one structured transport notice.
- Execution now rejects static semantic request violations before constructing execution bindings, returning the same ordered `CommandError.problems[]` core as request doctoring instead of a pre-mutation `WorkbookResult` failure, including distinct authored violations that happen to share the same message.
- Last-resort CLI failures now use the explicit `CLI_RUNTIME` problem stage before workbook execution, and an unwritable primary stdout stream terminates with a nonzero exit rather than appending a second diagnostic payload.
- `CommandError` deserialization now requires its fixed `status="REJECTED"` value instead of silently accepting contradictory or absent status input.
- Unexpected failures raised after workbook execution begins now retain the `WorkbookResult` envelope with `status=FAILED`, including the requested `SAVE_AS` or `OVERWRITE` intent as `NotWritten`, instead of being mislabeled as pre-execution `CommandError` rejections.
- An unavailable stdout transport no longer moves a primary diagnostic to stderr; GridGrind exits nonzero rather than emitting a competing result schema on the transport channel.
- `INTERNAL_ERROR` problems now use the canonical public title in their messages and causes instead of exposing arbitrary runtime exception text.
- Constructor-level and cross-fragment request diagnostics now resolve `PATH_BYTE_OFFSET` to the exact authored member or array token named by their JSON path; when no authored token exists, diagnostics retain the path without fabricating a byte coordinate.
- Request structural analysis now validates every authored occurrence of a known field, so a duplicate-key finding no longer hides an independently provable malformed, explicit-null, unknown-type, or nested-step problem in the repeated value.
- Duplicate root fields no longer leave an arbitrary first value available as a bound request fragment: the affected field, including `steps`, remains explicitly unbound while unrelated valid siblings remain available for analysis.
- Request-read diagnostics now anchor every object-member finding, including malformed scalar, explicit-null, unsupported-enum, and unknown-discriminator cases, at the property's opening quote; array-element and root-value findings retain their value-token offsets.
- Request intake now rejects trailing object and array commas, reports root-level non-JSON whitespace at byte zero, and classifies explicit `null` protocol versions and union discriminators with the same omission-only message used by every other request field.
- Request intake now distinguishes malformed UTF-8 from syntactically invalid JSON: non-UTF-8 request bytes fail as `INVALID_ENCODING`, while valid UTF-8 JSON syntax failures retain their existing `INVALID_JSON` classification and source coordinates.
- Request analysis now owns typed conversion from one tolerant parse: a syntax-faulted subtree cannot bind as valid, constructor-invalid fragments are retained as explicit intake findings, unaffected siblings remain available for analysis, and execution and doctoring share the same ordered problem collection instead of re-decoding the raw request. When a union discriminator is malformed, unknown fields that no variant permits are still reported; every valid authored discriminator branch contributes its independently provable shape findings without re-scanning raw members for each branch.
- Normal execution no longer drops independent structural request findings after the first one: it now returns the same deterministic problem collection as request doctoring before any workbook work starts, preserving exact value offsets and duplicate-key metadata rather than reconstructing a singular legacy location.
- Doctoring no longer rewrites malformed authored steps into synthetic placeholders to continue analysis, so every reported structural defect remains tied to the original request and no fabricated plan data can affect later diagnostics.
- Shadow packaging now feeds every `META-INF/services/**` descriptor to the service-file merger and rejects duplicate final JAR entries, preventing ServiceLoader registrations from being silently lost during distribution assembly.
- Conditional-formatting style write and readback no longer collapse theme, indexed, or tint color references to an RGB-only side channel; the documented examples and promoted request-fuzz fixtures now use the same tagged `ColorInput` shape enforced by the protocol.
- Response-file delivery now stages each payload beside its final destination, so a failed write leaves no partial response file and never replaces an existing one. When delivery falls back to stdout, GridGrind preserves the already-rendered, secret-safe primary payload unchanged and emits only one transport-only stderr notice rather than fabricating a secondary transport problem or changing the result status.
- Problems and warnings now use a total deterministic order across pipeline phase, source position, step, duplicate occurrence, code, and a nonserialized internal phase-local tie-breaker.

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
- Request intake now rejects malformed mismatched JSON container closers without stalling, accepts only RFC JSON whitespace, and classifies unrepresentable numeric and temporal scalar values at their authored path and UTF-8 byte offset before typed binding; late payload failures also retain their owned request location in CLI diagnostics.
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

[Unreleased]: https://github.com/resoltico/GridGrind/compare/v0.74.0...HEAD
[0.74.0]: https://github.com/resoltico/GridGrind/compare/v0.73.0...v0.74.0
[0.73.0]: https://github.com/resoltico/GridGrind/compare/v0.72.0...v0.73.0
[0.72.0]: https://github.com/resoltico/GridGrind/compare/v0.71.0...v0.72.0
[0.71.0]: https://github.com/resoltico/GridGrind/compare/v0.70.0...v0.71.0
[0.70.0]: https://github.com/resoltico/GridGrind/compare/v0.69.0...v0.70.0
[0.69.0]: https://github.com/resoltico/GridGrind/compare/v0.68.0...v0.69.0
