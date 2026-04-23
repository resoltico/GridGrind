---
afad: "3.5"
version: "0.55.0"
domain: DEVELOPER_JAZZER
updated: "2026-04-17"
route:
  keywords: [gridgrind, jazzer, fuzz, fuzzing, developer, local-only, regression, corpus, replay, promote, telemetry, composite-build, gradle, junit, xlsx, architecture]
  questions: ["how does jazzer fit into gridgrind", "where does jazzer live in this repo", "how is jazzer wired into the project", "what commands exist for jazzer", "where do jazzer corpus files and summaries go", "how do replay and promotion work", "what does jazzer cover in gridgrind"]
---

# Jazzer Developer Reference

**Purpose**: Authoritative architecture and policy reference for the GridGrind Jazzer layer.
**Companion references**:
- [DEVELOPER_JAVA.md](./DEVELOPER_JAVA.md) for the required shell-level Java 26 setup and the
  `./gradlew`-over-Brew-Gradle rule.
- [DEVELOPER_JAZZER_OPERATIONS.md](./DEVELOPER_JAZZER_OPERATIONS.md) for command usage, run
  lifecycle, replay, promotion, and cleanup.
- [DEVELOPER_JAZZER_COVERAGE.md](./DEVELOPER_JAZZER_COVERAGE.md) for the harness matrix, promoted
  seed inventory, and current fuzzing gaps.
- [DEVELOPER_GRADLE.md](./DEVELOPER_GRADLE.md) for the shared included build logic, root version
  catalog, and the reasons Jazzer stays a nested composite build.
- [DEVELOPER.md](./DEVELOPER.md) for the main GridGrind build, module architecture, and root
  quality gates.

---

## Status

This document is the canonical source of truth for Jazzer in GridGrind.

As of `2026-04-07`, the Jazzer layer is implemented and working end to end.

Implemented now:
- standalone nested Gradle build at `jazzer/`
- local-only wrapper-script entrypoints under `jazzer/bin/`
- deterministic nested-build support tests under `jazzer/src/test/java/`
- four active fuzz harnesses plus a regression-mode replay run
- local full-gate orchestration through root `./check.sh`, which runs root verification and
  nested Jazzer `check` sequentially
- per-target run history, latest-summary artifacts, and per-harness telemetry
- replay and promotion tooling
- pure-Java deterministic replay for the structured binary harnesses, so committed-seed replay
  and `jazzer/bin/replay` do not depend on Jazzer's native replay provider loading
- replay-expectation verification for every promoted metadata entry
- promoted-metadata self-validation for target-key consistency and referenced replay artifacts
- promoted-metadata refresh tooling for intentional generator and replay-shape changes
- committed custom seed floor made of promoted regression inputs and readable example requests
- style-aware `.xlsx` round-trip verification for formatting depth, hyperlink/comment metadata,
  named-range persistence, named-range normalization, data-validation persistence, table and
  autofilter persistence boundaries, and the explicit inspection-step pipeline
- local-only corpora, logs, finding artifacts, and cleanup commands
- semantic progress pulses and stall-aware outer-gate monitoring for long-running Jazzer stages
- in-flight support-test heartbeats during long deterministic support tests, so Stage 2 keeps
  reporting real work even when a single support case takes a while

Deliberately not implemented:
- any root-build dependency on Jazzer
- any GitHub CI wiring for Jazzer
- any committed local corpus files
- any committed libFuzzer dictionary files
- any `.xls`, `.xlsm`, macro, encryption, or signing fuzzing

If the implementation and these docs diverge, fix one or the other immediately.

---

## Core Decisions

GridGrind uses Jazzer as a separate, local-only fuzzing layer.

The non-negotiable decisions are:
- `jazzer/` is a nested build, not a root subproject.
- Root `./gradlew check`, root coverage, and GitHub CI stay independent of Jazzer.
- Root `./check.sh` is the supported local orchestrator for sequential root-plus-nested
  verification.
- Root project-file formatting excludes local-only instruction and scratch areas,
  so local workspace state cannot break the canonical quality gates.
- The root Gradle wrapper is reused; there is no second wrapper.
- Wrapper scripts are the supported operator surface, and active fuzz through them is explicitly
  `--no-daemon` plus wrapper-managed cleanup.
- GitHub Actions must never run active fuzzing; `JazzerHarnessRunner` hard-blocks it when
  `GITHUB_ACTIONS=true`.
- Generated state stays under `jazzer/.local/`.
- Only intentionally promoted regression inputs belong in versioned source.
- Workbook scope remains `.xlsx` only.

The deterministic JUnit suite remains the primary correctness contract. Jazzer extends local
exploration and bug discovery; it does not replace standard testing.

---

## Custom Seed Strategy

GridGrind now keeps a small committed custom seed floor for every replayable Jazzer harness.

Seed policy:
- the committed seed floor must stay intentionally curated and small
- every committed seed must represent a meaningfully different semantic shape, not just a new hash
- the local generated corpus remains the large exploratory pool and stays uncommitted
- a promoted seed is not automatically a deterministic ordinary test; promote first, then decide
  whether it also deserves a main-suite regression test
- deterministic replay for promoted binary seeds is defined by GridGrind's own replay seam, not by
  direct dependence on Jazzer's native `withJavaData(...)` bootstrap path
- every promoted metadata entry must carry a stable replay expectation and must remain replay-
  verified by deterministic support tests
- every promoted metadata entry must keep a matching harness target key and must continue to point
  at both an existing committed input and an existing replay-text artifact
- the promotion contract is bidirectional and enforced by `PromotionMetadataTest`: every file
  committed to a harness input directory must have a corresponding promoted-metadata entry, and
  every promoted-metadata entry must replay to its recorded expectation; hand-dropping a file
  into the input directory without running `jazzer/bin/promote` will fail the test suite

Target-specific strategy:
- `protocol-request` favors human-readable `.json` seeds promoted from public example requests
- readable request seeds should follow the current public example contract exactly, including the
  current `steps` shape
- readable request seeds should omit fields that intentionally default when omission is the public
  contract, so replay guards the real defaulting behavior instead of only an explicit equivalent
  payload
- readable request seeds should also track newly expanded factual read and mutation contracts when
  public examples add them, so richer shapes such as workbook protection, rich comments, advanced
  print setup, structured style colors or gradients, autofilter criteria or sort state, advanced
  table metadata, and formula-defined named ranges stay replayed end to end
- `protocol-workflow`, `engine-command-sequence`, and `xlsx-roundtrip` favor replay-verified
  binary seeds promoted from local corpus entries
- the meaning of a promoted binary seed is defined by its replay expectation and replay details,
  not by its filename alone
- engine command seeds may be reused for `.xlsx` round-trip seeds when replay confirms they persist
  cleanly in the round-trip harness
- the `.xlsx` round-trip verifier is expected to derive sheet-layout, style, metadata, and table
  expectations from the real pre-save workbook state; command replay exists only to bound which
  cell snapshots are compared after reopen

The operator goal is not to maximize seed count. The goal is to preserve a stable minimum floor of:
- representative success cases
- representative expected-invalid cases
- representative feature-family coverage

The current per-harness seed counts and the exhaustive committed filename inventory live in
[DEVELOPER_JAZZER_COVERAGE.md](./DEVELOPER_JAZZER_COVERAGE.md). Keep that file authoritative
instead of duplicating seed counts across multiple docs.

---

## Dictionary Promotion Policy

GridGrind does not currently ship any committed libFuzzer dictionary files.

Recommended-dictionary output from a single fuzz run is advisory only. It is never promoted
automatically.

Promotion gate:
- the recommendation must repeat across multiple long runs of the same harness
- the candidate entries must be human-recognizable and harness-relevant
- the candidate dictionary must stay small and harness-specific
- the candidate dictionary must produce a measurable benefit in coverage or execution efficiency
  when compared with the same harness run without the dictionary

Disqualifiers:
- entries that appear to be one-run noise or corpus-specific overfitting
- very large opaque dictionaries assembled directly from one session's recommendation block
- dictionaries shared across unrelated harnesses
- dictionaries that improve only corpus size growth without improving coverage, feature count, or
  useful exploration speed

Promotion standard:
- require repeated appearance across at least three independent long runs of the same harness
- require a curated dictionary file containing only the smallest useful subset of entries
- require an A/B comparison for that harness using comparable duration and settings
- require documentation updates describing why the dictionary exists and what evidence justified it

If GridGrind ever adopts committed dictionaries, they must:
- be stored as harness-specific files under a dedicated `jazzer/` subdirectory
- remain local-only tooling for the nested Jazzer build
- never become an implicit dependency of the root build or CI

Until those criteria are met, the correct action for recommended-dictionary output is:
- keep the log
- continue collecting evidence
- do not commit a dictionary

---

## Scope Lock

In scope:
- `.xlsx` only
- protocol JSON parsing and validation
- ordered `DefaultGridGrindRequestExecutor` workflows
- direct workbook-command execution
- `.xlsx` save and reopen invariants

Out of scope:
- `.xls`
- `.xlsm`
- macros or VBA
- encryption and signing
- non-Excel formats

---

## Repository Topology

Top-level placement:

```text
GridGrind/
├── docs/
│   ├── DEVELOPER.md
│   ├── DEVELOPER_GRADLE.md
│   ├── DEVELOPER_JAZZER.md
│   ├── DEVELOPER_JAZZER_OPERATIONS.md
│   └── DEVELOPER_JAZZER_COVERAGE.md
├── jazzer/
│   ├── README.md
│   ├── settings.gradle.kts
│   ├── build.gradle.kts
│   ├── gradle.properties
│   ├── bin/
│   │   ├── _run-task
│   │   ├── regression
│   │   ├── refresh-promoted-metadata
│   │   ├── fuzz-protocol-request
│   │   ├── fuzz-protocol-workflow
│   │   ├── fuzz-engine-command-sequence
│   │   ├── fuzz-xlsx-roundtrip
│   │   ├── fuzz-all
│   │   ├── status
│   │   ├── report
│   │   ├── list-findings
│   │   ├── list-corpus
│   │   ├── replay
│   │   ├── promote
│   │   ├── clean-local-findings
│   │   └── clean-local-corpus
│   ├── src/
│   │   ├── main/
│   │   │   └── java/
│   │   │       └── dev/erst/gridgrind/jazzer/
│   │   │           ├── support/
│   │   │           │   ├── GridGrindFuzzData.java
│   │   │           │   ├── JazzerGridGrindFuzzData.java
│   │   │           │   └── ReplayGridGrindFuzzData.java
│   │   │           └── tool/
│   │   │               └── JazzerRegressionRunner.java
│   │   ├── test/
│   │   │   └── java/
│   │   │       └── dev/erst/gridgrind/jazzer/
│   │   │           ├── support/
│   │   │           └── tool/
│   │   └── fuzz/
│   │       ├── java/
│   │       │   └── dev/erst/gridgrind/jazzer/
│   │       │       ├── engine/
│   │       │       └── protocol/
│   │       └── resources/
│   │           └── dev/erst/gridgrind/jazzer/
│   │               ├── engine/
│   │               ├── protocol/
│   │               └── promoted-metadata/
│   └── .local/
│       ├── run-lock/
│       └── runs/
│           ├── regression/
│           ├── protocol-request/
│           ├── protocol-workflow/
│           ├── engine-command-sequence/
│           └── xlsx-roundtrip/
└── ...
```

Tree rules:
- `jazzer/` is the only top-level home for Jazzer code and Jazzer-local state.
- `jazzer/src/main/java` contains shared support, telemetry, replay, report, and promotion code.
- `jazzer/src/main/java/dev/erst/gridgrind/jazzer/support/GridGrindFuzzData` is the only allowed
  structured-generator seam for scalar fuzz-data consumption; deterministic replay must stay on
  the project-owned pure-Java implementation rather than calling Jazzer native replay APIs
  directly.
- `jazzer/src/test/java` contains deterministic tests for the Jazzer support and reporting layer.
- `protocol-request` and `protocol-workflow` remain harness names for the canonical JSON request
  surface; they are not references to a live top-level `protocol` Gradle module. The product
  module split is `contract` plus `executor`, while the Jazzer harness names intentionally stay
  stable so historical findings, corpora, and promoted metadata paths remain coherent.
- `jazzer/src/fuzz/java` contains only fuzz harness classes.
- `jazzer/src/fuzz/resources` contains only committed regression inputs and promotion metadata.
- `jazzer/.local/` is the only approved location for generated local fuzz state.

---

## Build Model

`jazzer/settings.gradle.kts` uses `includeBuild("..")`, so the nested build consumes the live
local `engine` and `protocol` modules without publishing snapshots. It also imports the shared root
version catalog from `../gradle/libs.versions.toml` and resolves `gridgrind.*` plugins from the
shared included build logic under `../gradle/build-logic`.

`jazzer/build.gradle.kts` is intentionally thin and now delegates its behavior to the
`gridgrind.jazzer-conventions` plugin. The nested build still provides:
- standard `main` source set for shared support and local operator tooling
- standard `test` source set for deterministic support tests
- custom `fuzz` source set for Jazzer harness classes and committed regression inputs
- shared version-catalog dependency alignment for JUnit 6, Jazzer 0.30.0, Jackson 3.1.2,
  Apache POI 5.5.1, Log4j, and the local product modules
- `JavaExec` tasks for local operator commands and per-harness Jazzer execution
- one explicit JUnit Platform launcher entrypoint for running Jazzer harnesses outside Gradle's
  `Test` result pipeline
- one aggregate `jazzerRegression` lifecycle task over the four per-harness regression tasks
- one dedicated metadata-refresh task for promoted replay artifacts
- local cleanup tasks
- shared Spotless and PMD enforcement inherited from the main build logic
- a Jazzer-specific PMD profile for support and operator code plus a fuzz-harness PMD profile for
  `@FuzzTest` entrypoints
- a dedicated `jazzerCoverageVerification` gate for deterministic support-contract classes,
  separate from the root modules' blanket 100% coverage rule
- nested-build `check` coverage for deterministic Jazzer support tests plus regression replay only

Single source of truth:
- `jazzer/src/main/resources/dev/erst/gridgrind/jazzer/support/jazzer-topology.json` is the
  authoritative registry for Jazzer harness keys, class names, fuzz-method names, run-target task
  names, and working directories
- runtime helpers such as `JazzerHarness` and `JazzerRunTarget` now load that topology instead of
  hardcoding duplicate enum metadata
- the nested Gradle build consumes the same topology when registering fuzz and regression tasks, so
  build task names and runtime replay metadata cannot drift independently

Important boundary:
- the root build must not include `jazzer/`
- the nested build may depend on root modules through the composite build
- the nested build may evolve independently without affecting root CI

Nested-build verification model:
- `./gradlew --project-dir jazzer test` runs deterministic support tests only
- `./gradlew --project-dir jazzer jazzerRegression` replays the committed seed floor through four
  isolated per-harness regression tasks
- `./gradlew --project-dir jazzer check` runs Spotless, PMD, deterministic support tests,
  `jazzerCoverageVerification`, and `jazzerRegression`
- root `./check.sh` runs root-project verification first, then nested Jazzer `check`, then
  release-packaging smoke verification
- active fuzz tasks remain explicit opt-in local commands and are never dependencies of nested-build
  `check`
- active fuzz scripts launch one harness class at a time through `JazzerHarnessRunner`, which
  enforces the repository contract of exactly one `@FuzzTest` per harness class and then delegates
  to Jazzer's own command-line `JUnitRunner` so internal command-line mode, `jazzer.max_duration`,
  and `jazzer.max_executions` are honored consistently
- active fuzz wrappers force `--no-daemon` and tear down the launched Gradle client tree on
  interrupt or timeout so the supported local surface does not leave a harness JVM behind
- those wrapper scripts must also remain compatible with stock macOS `/bin/bash` 3.2 under
  `set -u`, even when no optional Gradle arguments are supplied
- active fuzz harness execution hard-fails when `GITHUB_ACTIONS=true`, so GitHub remains a
  deterministic-only verification surface even if someone wires an active fuzz task into a
  workflow by mistake
- active fuzz launcher JVMs preload the project-owned `JazzerPremainAgent`, which publishes
  startup instrumentation to Byte Buddy before Jazzer's JUnit extension runs and avoids fragile
  late attach behavior on Java 26

Execution discipline:
- do not run root Gradle builds and nested Jazzer Gradle builds in parallel
- run root verification and Jazzer verification sequentially
- shared workspace outputs and caches can otherwise produce stale classpath or runtime state

---

## Run Targets

| Target | Gradle Task | Mode | Working Directory |
|:-------|:------------|:-----|:------------------|
| `regression` | `jazzerRegression` | regression replay | `jazzer/.local/runs/regression/` |
| `protocol-request` | `fuzzProtocolRequest` | active fuzzing | `jazzer/.local/runs/protocol-request/` |
| `protocol-workflow` | `fuzzProtocolWorkflow` | active fuzzing | `jazzer/.local/runs/protocol-workflow/` |
| `engine-command-sequence` | `fuzzEngineCommandSequence` | active fuzzing | `jazzer/.local/runs/engine-command-sequence/` |
| `xlsx-roundtrip` | `fuzzXlsxRoundTrip` | active fuzzing | `jazzer/.local/runs/xlsx-roundtrip/` |

Operator rule:
- the scripts under `jazzer/bin/` are the supported entrypoints
- raw `./gradlew --project-dir jazzer ...` remains available only for deterministic nested-build
  verification; active fuzzing is supported only through `jazzer/bin/*`

`jazzer/bin/fuzz-all` is intentionally a shell sequencer over the four per-harness scripts so
each harness still gets its own lock, run history, summary, and telemetry artifacts.

`jazzer/bin/regression` follows the same isolation principle inside the nested build:
- one public command
- four isolated per-harness regression launcher tasks
- one aggregate `regression` summary under `jazzer/.local/runs/regression/`

This design is intentional. The regression seed floor is replayed one harness at a time so the
Jazzer build never depends on Gradle's `Test` binary-results store for fuzz-harness execution.
Deterministic support tests still run through the standard Gradle `test` task.

---

## Storage Contract

Committed and reviewable:
- harness classes under `jazzer/src/fuzz/java/`
- shared support and operator tooling under `jazzer/src/main/java/`
- promoted custom seeds under `jazzer/src/fuzz/resources/.../*Inputs/...`
- promotion metadata under `jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata/`
- documentation under `docs/` and `jazzer/README.md`

Local-only, never committed:
- `jazzer/.local/run-lock/`
- `jazzer/.local/runs/*/.cifuzz-corpus/`
- `jazzer/.local/runs/*/latest.log`
- `jazzer/.local/runs/*/latest-summary.json`
- `jazzer/.local/runs/*/latest-summary.txt`
- `jazzer/.local/runs/*/telemetry/*.json`
- `jazzer/.local/runs/*/findings/*`
- `jazzer/.local/runs/*/history/*`

Cleanup contract:
- `jazzer/bin/clean-local-corpus` removes generated corpus trees only
- `jazzer/bin/clean-local-findings` removes non-corpus local run state, including logs, summaries,
  telemetry, and replayed finding artifacts

---

## Reporting and Replay Model

Every supported fuzz run now leaves behind:
- a per-run `history/<timestamp>/run.log`
- a per-run `history/<timestamp>/summary.json`
- a per-run `history/<timestamp>/summary.txt`
- `latest-summary.json` and `latest-summary.txt` in the run directory
- per-harness telemetry snapshots under `telemetry/`

The operator tooling adds:
- `status` for a concise latest-summary table across targets
- `report` for detailed latest-summary output
- `list-corpus` for corpus state plus committed promoted inputs
- `list-findings` for replayed local finding artifacts
- `replay` for re-running one local input against one harness
- `promote` for copying one input into committed regression resources and writing metadata

Finding flow:
1. active fuzzing writes raw local finding artifacts such as `crash-*`
2. summary generation replays each raw artifact and writes human-readable finding artifacts
3. `replay` lets an operator inspect one input explicitly
4. `promote` copies a selected input into committed regression resources
5. `regression` replays the committed inputs on every local regression run
6. confirmed product bugs should still graduate into deterministic tests in the main suite

Summary semantics:
- `status` and `report` count only active findings as findings
- expected-invalid artifacts are surfaced separately when an old raw crash file now replays inside
  the documented invalid-input contract
- replay-clean artifacts are still surfaced separately so older local crash files that now replay
  cleanly are visible without being misreported as active failures

Telemetry semantics:
- sequence telemetry records operation kinds or command kinds per harness
- style telemetry records formatting-family usage such as font height, fill, font color,
  underline, strikeout, and border shapes
- protocol workflow telemetry also records source-type and persistence-type coverage
- response and error telemetry remain the primary triage signal for unexpected outcomes

---

## Compatibility Notes and Caveats

Jazzer/JUnit:
- Jazzer 0.30.0 publishes JUnit integration in "JUnit 5" terminology
- the nested build pins JUnit 6 explicitly and works with Gradle 9.4.1 in practice

Structured replay:
- replay for the structured harnesses uses Jazzer's internal
  `com.code_intelligence.jazzer.driver.FuzzedDataProviderImpl.withJavaData(byte[])`
  so replay semantics match the live harness data-provider semantics
- Jazzer 0.30.0's published jars omit `com.code_intelligence.jazzer.utils.UnsafeProvider`, even
  though `FuzzedDataProviderImpl` depends on it
- the nested build therefore vendors a minimal compatible copy of that single helper under
  `jazzer/src/main/java/com/code_intelligence/jazzer/utils/UnsafeProvider.java`
- this vendored helper exists only to make structured replay work in the isolated local Jazzer
  layer

Expected warnings:
- active fuzzing and structured replay still emit `sun.misc.Unsafe` deprecation warnings on
  Java 26
- these warnings are expected for Jazzer 0.30.0 and do not indicate a GridGrind defect

Configuration cache:
- active fuzz tasks participate in Gradle configuration caching
- local JavaExec operator tasks intentionally run with `--no-configuration-cache` via the wrapper
  script because they are local-only utilities and do not benefit enough from cache participation
  to justify the extra Gradle noise

Operator surfaces:
- `./gradlew --project-dir jazzer test` and `./gradlew --project-dir jazzer check` are the
  deterministic nested-build verification entrypoints
- `jazzer/bin/*` is the supported live operator surface for active fuzzing, replay, promotion,
  status, report, and cleanup
- raw Gradle active-fuzz tasks are not a supported operator method and are intentionally excluded
  from the documented fuzz workflow

---

## Cross-References

Use:
- [DEVELOPER_JAZZER_OPERATIONS.md](./DEVELOPER_JAZZER_OPERATIONS.md) when the question is
  "how do I run or inspect Jazzer?"
- [DEVELOPER_JAZZER_COVERAGE.md](./DEVELOPER_JAZZER_COVERAGE.md) when the question is
  "what does Jazzer currently cover, and what is still missing?"

Keep this file focused on architecture, boundaries, storage, and policy.
