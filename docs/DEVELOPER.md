---
afad: "4.0"
version: "0.62.0"
domain: DEVELOPER
updated: "2026-05-01"
route:
  keywords: [gridgrind, build, gradle, architecture, coverage, jacoco, pmd, errorprone, spotless, java26, devcontainer, zulu26, engine, contract, executor, authoring-java, cli]
  questions: ["how do I build gridgrind", "how do I run tests", "what is the preferred contributor setup for gridgrind", "what is the gridgrind architecture", "how are quality gates configured", "what are the coverage requirements"]
---

# Developer Reference

**Purpose**: Build, test, architecture, and quality gate reference for GridGrind contributors.
**Preferred contributor path**: the committed devcontainer in
[DEVELOPER_DEVCONTAINER.md](./DEVELOPER_DEVCONTAINER.md), entered either through VS Code Dev
Containers or the tooling-agnostic Dev Container CLI.
**Fallback contributor path**: host-native Java 26 plus Docker Desktop as documented in
[DEVELOPER_JAVA.md](./DEVELOPER_JAVA.md) and [DEVELOPER_DOCKER.md](./DEVELOPER_DOCKER.md).

Companion references:
- [DEVELOPER_DEVCONTAINER.md](./DEVELOPER_DEVCONTAINER.md)
- [DEVELOPER_DOCKER.md](./DEVELOPER_DOCKER.md)
- [DEVELOPER_JAVA.md](./DEVELOPER_JAVA.md)
- [DEVELOPER_GRADLE.md](./DEVELOPER_GRADLE.md)
- [JAVA_AUTHORING.md](./JAVA_AUTHORING.md)
- [DEVELOPER_CONTRACT_REPLACEMENT_ADR.md](./DEVELOPER_CONTRACT_REPLACEMENT_ADR.md)
- [DEVELOPER_JAZZER.md](./DEVELOPER_JAZZER.md)
- [DEVELOPER_JAZZER_OPERATIONS.md](./DEVELOPER_JAZZER_OPERATIONS.md)
- [DEVELOPER_JAZZER_COVERAGE.md](./DEVELOPER_JAZZER_COVERAGE.md)

---

## Preferred Contributor Workflow

GridGrind's preferred local contributor workflow is the committed devcontainer:

- keep the repository on the local macOS filesystem
- materialize the committed devcontainer through a devcontainer-spec client
- use VS Code `Reopen in Container` for the most integrated editor path, or
  `devcontainer up` plus `devcontainer exec` for the tooling-agnostic path
- run Java, Gradle, Jazzer, shell verification, and Docker-backed repo gates from the container
  terminal

That path keeps the contributor Java stack aligned with CI's Azul Zulu 26 vendor while isolating
the repository from host-side Java extension daemons and attach probes. Host-native Java 26 work
remains supported, but the devcontainer is the documented default.

---

## Architecture

GridGrind is a six-module Gradle project with compiler-enforced JPMS boundaries:

```
excel-foundation/
            Shared Excel-domain foundation. Owns the POI-free `.xlsx` enums,
            limits, and span/value objects shared by contract, engine,
            executor-facing tests, and downstream adapters.

engine/     Apache-POI-backed workbook engine. Owns mutable workbook state,
            workbook mutation rules, and factual workbook inspection.

contract/   Canonical GridGrind contract model. Owns the public request and
            response types, catalog metadata, JSON codecs, and transport-
            neutral contract shapes.

executor/   The only execution bridge from the canonical contract into the
            workbook engine. Owns plan validation, request execution, and
            parity-facing integration.

authoring-java/
            Fluent Java authoring layer. Owns selector builders, authored
            step builders, and canonical `WorkbookPlan` emission without
            exposing engine internals or pulling executor runtime concerns
            into the module boundary.

cli/        Thin transport adapter. Reads a contract request from stdin
            or --request file, delegates to executor, writes the response,
            and exposes help/template/catalog discovery commands.
```

The CLI is not the core. The foundation is `excel-foundation` plus `engine` plus `contract` plus
`executor`. The `authoring-java` API and the CLI are two downstream surfaces on top of that
foundation. Future adapters (HTTP, gRPC, library embedding) can be added without touching
`excel-foundation`, `engine`, `contract`, or `executor`.

The enforced dependency graph is:

```text
dev.erst.gridgrind.authoring -> dev.erst.gridgrind.contract -> dev.erst.gridgrind.excel.foundation
dev.erst.gridgrind.cli -> dev.erst.gridgrind.executor -> dev.erst.gridgrind.contract -> dev.erst.gridgrind.excel.foundation
dev.erst.gridgrind.executor -> dev.erst.gridgrind.engine -> dev.erst.gridgrind.excel.foundation
```

`authoring-java` and `cli` do not depend on `engine`, `authoring-java` also does not depend on
`executor`, and `executor` is the only module allowed to bridge from the canonical contract into
workbook execution. `excel-foundation` is the only shared Excel-domain surface allowed to sit
below both `contract` and `engine`, so the public contract no longer imports POI-backed engine
internals just to reuse enums or limits. Shared Java build conventions enable
`modularity.inferModulePath`, so the `module-info.java` descriptors in all six product modules
participate in normal local builds, CI, and release verification.

The highest-churn architecture seams are intentionally split too:
- `GridGrindProtocolCatalog` owns the top-level catalog assembly, while `GridGrindProtocolCatalogFieldGroupSupport` owns the large nested/plain field-shape descriptor registry and `GridGrindProtocolCatalogLookupSupport` owns lookup/search behavior.
- Jazzer request generation is no longer one monolith: the `OperationSequence*` family now splits orchestration, selector helpers, mutation/inspection/command factories, observation-side assertions, bounded payload generation, and chart payload generation into focused seams, while `WorkbookInvariantChecks` is split across workbook/cell/engine-surface invariant helpers.
- Build-failing architecture audits protect these boundaries. Contract source must stay on `excel-foundation`, direct `Cell.setCellFormula(...)` calls must stay inside `ExcelFormulaWriteSupport`, POI reflective access must stay inside `PoiPrivateAccessSupport`, and the split hotspot files must stay below their enforced size ceilings. `PoiPrivateAccessCompatibilityTest` also fails fast if a POI upgrade stops resolving the registered private sheet-clone, picture-catalog, fill-registry, or relation-removal seams that the engine still depends on.

## Contract Replacement Mode

The legacy monolithic `protocol` module is gone. GridGrind is now in hard-break contract-replacement
mode: new top-level contract surface growth must happen through the `contract` plus `executor`
split and the accepted post-replacement architecture, not by reintroducing monolithic
transport-and-execution ownership. The accepted architecture decision record for that freeze is
[DEVELOPER_CONTRACT_REPLACEMENT_ADR.md](./DEVELOPER_CONTRACT_REPLACEMENT_ADR.md).

---

## Foundations

| Component | Version |
|:----------|:--------|
| Contributor devcontainer | pinned glibc base image plus Azul Zulu 26 JDK; see [DEVELOPER_DEVCONTAINER.md](./DEVELOPER_DEVCONTAINER.md) |
| Java | 26 |
| Docker runtime | Docker Desktop daemon plus `docker buildx` reachable through the active shell `docker` command; smoke and release verification use an anonymous `DOCKER_CONFIG` while still targeting the active local Docker engine |
| Apache POI | 5.5.1 |
| Jackson Databind | 3.1.2 |
| JUnit Jupiter | 6.1.0-RC1 |
| Log4j Core | 2.25.3 |

GridGrind's runtime, product modules, and shared included build logic under `gradle/build-logic`
all target Java 26 now. The included build is no longer a JVM 25 exception: it compiles with
Kotlin `2.4.0-Beta2`, emits JVM 26 bytecode directly, and stays aligned with the repository's
single Java baseline instead of carrying a separate bytecode-level footnote.

Jackson dependency note: Jackson 3.x databind intentionally still uses the
`com.fasterxml.jackson.core:jackson-annotations` artifact and package namespace. That is upstream
Jackson design, not a GridGrind version-skew bug.

---

## Commands

`./gradlew check` remains the root-project code-quality gate: Spotless formatting,
explicit-import verification, Error Prone, PMD, tests, and JaCoCo coverage verification for
`engine`, `contract`, `executor`, `authoring-java`, and `cli`. `./check.sh` is the supported
whole-repo deterministic gate: root `check` plus `coverage`, nested Jazzer `check`,
`:cli:shadowJar`, architecture-split shell regressions, packaged-JAR CLI contract verification,
release-surface shell checks, and a Docker smoke test that runs the image from a non-default
working directory with weird request/response/save paths while also asserting the published
help/version/response contract semantically. `check.sh`
intentionally lives in the repository root as the canonical contributor entrypoint, while
`scripts/` contains helper scripts invoked by the root gate and GitHub workflows. Both should be
clean before a release-quality change is considered done. When Stage 1 reaches long-running Gradle
`Test` tasks, the shared test conventions emit
`[GRADLE-TEST-PULSE]` lines with class-start, class-complete, and throttled test-progress facts
so the local watchdog tracks semantic execution rather than mistaking quiet test output for a
hang. The fixed five-stage inventory and the Stage 4 shell-regression list are canonically owned
by `scripts/check-stage-contract.sh`, so root-gate usage text and execution wiring stay aligned as
release-surface shell coverage evolves.

Use `./gradlew`, not path-global `gradle`. GridGrind's CLI, fat JAR, release flow, and
`./check.sh` all depend on the ambient shell `java`, so the local shell must resolve Java 26 from
a real full JDK. Keep the repository on the local filesystem when you run the full verification
loop; mounted external volumes are outside the supported setup because Gradle project-cache and
JaCoCo file locking can fail there on macOS. If you edit from a mounted volume, run the heavy
build/test loop from a local-disk worktree or disposable mirror instead. `./check.sh` now fails
fast if the shell runtime is wrong. The same root gate also forces `--no-daemon`, isolates
`GRADLE_USER_HOME` to a repo-scoped path under `tmp/gradle-user-home` by default, and shares one
repo-wide verification lock with `scripts/validate-devcontainer.sh`, `scripts/docker-smoke.sh`,
and `jazzer/bin/*` so top-level verification entrypoints cannot overlap accidentally.
Docker smoke and release verification should likewise stay independent from personal Docker login
state by using an anonymous `DOCKER_CONFIG` while still targeting the active local Docker engine,
and local Docker verification now requires `docker buildx` because Stage 5 builds through
`docker buildx build --load` instead of Docker's legacy builder path.
The committed contributor devcontainer is validated separately through
`./scripts/validate-devcontainer.sh`, which builds the contributor image, checks its pinned Java
and tooling surface, guards the portable devcontainer contract, and verifies the committed VS Code
overlay contract in `.devcontainer/devcontainer.json`.

```bash
# Run the local full-stack gate
./check.sh

# Run the root-project CI gate
./gradlew check

# Targeted iteration during development
./gradlew test                   # tests only, no quality gates
./gradlew coverage               # tests + coverage verification + HTML/XML report generation
./gradlew spotlessCheck          # formatting check only
./gradlew spotlessApply          # auto-format in place (run before spotlessCheck)
./gradlew pmdMain pmdTest        # PMD only

# Build artifacts
./gradlew :cli:shadowJar
./gradlew :cli:run --args="--request examples/budget-request.json"
./gradlew :cli:run --args="--version"
./gradlew :cli:run --args="--print-request-template"
./gradlew :cli:run --args="--print-protocol-catalog"
./gradlew :cli:run --args="--print-example BUDGET"
./gradlew :cli:run --args="--print-task-catalog"
./gradlew :cli:run --args="--print-task-plan DASHBOARD"
./gradlew :cli:run --args='--print-goal-plan "monthly sales dashboard with charts"'
./scripts/docker-smoke.sh
./scripts/validate-devcontainer.sh
```

The protocol catalog is generated from the contract record signatures. If you change request
records, nested tagged unions, or plain input records, keep the field-shape output authoritative:
every field should still publish required/optional status and the exact nested/plain group
accepted by polymorphic inputs. Contract-bearing discovery prose such as low-memory mode limits,
formula parse boundaries, workbook-health summaries, and CLI help labels now live in the
contract-owned metadata layer instead of thin downstream string copies. `GridGrindContractText`
owns stable wording, `GridGrindExecutionModeMetadata` owns the structured EVENT_READ and
STREAMING_WRITE contract rules plus their validation messages, and `CliSurface` owns the help
section labels, key/value entries, flags, docs links, and example routing that `cli` renders.
`GridGrindTaskCatalog` owns the high-level task descriptors, `GridGrindTaskPlanner` derives
starter request scaffolds from those descriptors, and `GridGrindGoalPlanner` turns one freeform
goal into ranked task candidates without hardcoding scenario permutations. Request linting is
also part of that authoritative public surface now: `executor` owns `GridGrindRequestDoctor`, and
the packaged-artifact verifier exercises the emitted doctor report instead of relying only on
module-local tests. The build now also includes a contract-owned public-surface linter that fails
if docs, generated help, catalog summaries, shipped examples, or shared runtime diagnostics
mention a canonical mutation, assertion, or inspection id that is not registered in the catalog
vocabulary.
The catalog build path still uses a small internal
`contract.catalog.gather` seam for the two cases that genuinely benefit from Stream Gatherers:
ordered uniqueness and reflected field-metadata expansion. The built-in request template remains
an ordinary constant because no domain gatherer semantics are needed there. Keep
`--print-request-template` as the smallest valid machine-readable request; workflow snippets such
as the no-save `ANALYZE_WORKBOOK_FINDINGS` lint pass belong in help text, public docs, and the
example request set instead of the emitted template itself.

---

## GitHub Workflows

Release automation is split across three workflows:

- `CI` runs the supported whole-repo deterministic gate through `./check.sh`, then validates the
  committed contributor devcontainer surface so the preferred local development environment cannot
  silently drift away from the documented contract.
- `Release` builds the fat JAR, black-box verifies the packaged CLI discovery surface, publishes
  the GitHub Release idempotently when a `v*` tag is pushed, and verifies that the published
  release object and `gridgrind.jar` asset exist.
- `Container` first runs the same Docker smoke script against the local Dockerfile build, then
  builds and publishes the multi-arch GHCR image, verifies with an isolated anonymous Docker
  config that both the exact version tag and `latest` are publicly pullable and runnable, then
  black-box verifies the published CLI help, protocol catalog, task catalog, task planner,
  goal planner, and doctor surfaces from both tags before pruning older container package
  versions.
- `Gradle wrapper validation` runs when wrapper files change and validates the checked-in wrapper
  surface.

Targeted `workflow_dispatch` reruns for `Release` and `Container` are guarded deliberately: before
either workflow builds or publishes anything, the checked-out tag must still match
`gradle.properties`, resolve to the exact remote tag commit, remain reachable from `main`, and
already have successful `Check` plus `Docker smoke` runs on that exact commit. The container
workflow also publishes explicit OCI provenance and SBOM attestations for the released image.

The container cleanup step intentionally uses `gh api` against GitHub Packages instead of
`actions/delete-package-versions`. That action still runs on Node20, while GitHub is moving
JavaScript actions to Node24, and raw count-based cleanup is incorrect for GHCR multi-arch
images because one release produces several package-version records: the tagged OCI index plus
multiple untagged platform and attestation manifests.

The workflow therefore keeps the five newest tagged releases (`X.Y.Z`) as anchors and deletes
only package-version records older than the oldest retained anchor. This preserves complete
release groups instead of splitting a release across the retention boundary.

---

## Quality Gates

`./gradlew check` runs each of the following root-project gates. Any failure fails the build.
`./check.sh` runs these same root-project gates, then runs nested Jazzer verification, builds the
CLI fat JAR, syntax-checks the release-surface shell scripts, and finally runs the Docker smoke
script.
The release/publication shell verifiers and their regressions now use repo-local disposable
scratch under `tmp/` rather than depending on macOS `/var/folders` temp behavior, so local runs
and CI exercise the same deterministic shell surface.
If Docker or shell-script stages materialize temporary secret-bearing fixtures, those fixtures must
obey the same filesystem-security contract as production instead of weakening the runtime policy
just to make smoke tests pass.

### Error Prone

Compile-time static analysis. These checks are promoted to errors (build fails):

- `BadImport`
- `BoxedPrimitiveConstructor`
- `CheckReturnValue`
- `EqualsIncompatibleType`
- `JavaLangClash`
- `MissingCasesInEnumSwitch`
- `MissingOverride`
- `ReferenceEquality`
- `StringCaseLocaleUsage`

### PMD

Structural analysis. Two rulesets:

- `gradle/pmd/ruleset.xml` — production code: full `errorprone`, `bestpractices`, `design`,
  `multithreading`, `performance`, `security`, documentation (`CommentRequired` enforces Javadoc
  on all public types and methods).
- `gradle/pmd/test-ruleset.xml` — test code: same categories with relaxed assertion volume and
  method-count limits; `CommentRequired` enforces class-level Javadoc only.

### Spotless

Google Java Format 1.35.0. Removes unused imports. Run `./gradlew spotlessApply` to auto-format
before committing; `./gradlew spotlessCheck` (run by `check`) will fail if formatting is off.
Project-file formatting intentionally excludes local-only instruction and scratch areas,
so personal workspace state cannot destabilize the canonical quality gates.

### JaCoCo

100% line and branch coverage required across all modules.

Coverage-gate protocol:
- never rely on JaCoCo defaults for verification semantics
- per-module verification must set both `LINE` and `BRANCH` counters explicitly
- per-module reports and verification must read all local `build/jacoco/*.exec` files, not only
  `test.exec`
- aggregated root coverage must read all subproject `build/jacoco/*.exec` files as well

This rule is especially important in GridGrind because `executor` already has more than one local
`Test` task. A hardcoded `test.exec` assumption would silently exclude `parityTest` coverage from
the gate. See [DEVELOPER_GRADLE.md](./DEVELOPER_GRADLE.md) for the canonical build-logic
protocol.

After the `contract` plus `executor` split, `contract` coverage intentionally includes downstream
consumer execution data from `executor` and `cli` in addition to `contract`'s own tests. That is
not a loophole: it is the correct ownership model for a canonical public contract whose behavior
is exercised both by direct DTO/catalog tests and by the executor and CLI seams that consume it.

| Module | Line Coverage | Branch Coverage |
|:-------|:-------------|:----------------|
| `engine` | 100% | 100% |
| `contract` | 100% | 100% |
| `executor` | 100% | 100% |
| `authoring-java` | 100% | 100% |
| `cli` | 100% | 100% |

`./gradlew check` enforces the thresholds but does not write report files. Run
`./gradlew coverage` to enforce thresholds and generate HTML/XML reports at
`build/reports/jacoco/` in each module plus an aggregated cross-module report at
`build/reports/jacoco/aggregated/`.
The root coverage tasks discover participating JaCoCo-enabled Java subprojects dynamically and
collect every module `build/jacoco/*.exec` file, so adding another product module must not
require hand-editing a root hardcoded task list.

The table above applies only to the root product modules. The nested Jazzer build has a separate
coverage contract because its local-only generator, telemetry, and operator classes are exercised
primarily through regression replay and live fuzzing rather than ordinary unit tests. Run
`./gradlew --project-dir jazzer check` to enforce Jazzer's dedicated coverage scope together with
its shared Spotless and PMD gates. The entire extracted `OperationSequence*` generator family
follows the same coverage scope as `OperationSequenceModel`, and direct selector-sweep seam tests
still exercise those split generator layers so the exclusion does not mean "untested".

Supported `jazzer/bin/*` wrappers are part of that operator contract too. They must remain
compatible with stock macOS `/bin/bash` 3.2 under `set -u`; do not assume Bash 4+ empty-array
expansion semantics in shell wrappers that contributors run directly.

---

## Execution Semantics

Mutation steps run first in authored order, assertion and inspection steps run wherever they are
placed in authored order, and persistence happens only after every step succeeds.

If any step or calculation phase fails, the request returns a structured error and workbook
persistence is skipped. This gives agents deterministic failure semantics at the `.xlsx` output
boundary: GridGrind does not emit a partially written workbook file.

---

## Workflow Fixtures

The shipped JSON fixtures are generated from the contract-owned example registry in
checkout-rooted form. Refresh them with
[`scripts/sync-generated-examples.sh`](../scripts/sync-generated-examples.sh), or print any one of
the artifact-native built-in examples directly with `gridgrind --print-example <id>`. The full
map, path-rooting rules, and verification loop live in [EXAMPLES.md](./EXAMPLES.md). These
fixtures and authoring examples cover the core surface:

| File | What It Tests |
|:-----|:-------------|
| `examples/budget-request.json` | Range write, style, formula, workbook summary, cells, window, and schema inspection steps |
| `examples/file-hyperlink-health-request.json` | File-hyperlink authoring plus explicit hyperlink metadata and hyperlink-health analysis |
| `examples/introspection-analysis-request.json` | Inspection-heavy workbook showcasing factual reads plus targeted formula, hyperlink, named-range, and aggregate workbook analysis together |
| `examples/java-authoring-workflow.java` | Fluent Java authoring example showing selector-first table-key targeting, source-backed row selection, ordered inspection, and assertion steps without hand-written JSON |
| `examples/array-formula-request.json` | Dedicated array-formula authoring, array-group readback, and group clearing |
| `examples/custom-xml-request.json` | Existing-workbook custom-XML mapping discovery, XML export, file-backed XML import, and factual cell reread |
| `examples/signature-line-request.json` | Signature-line authoring with drawing-object readback and authored anchor replacement |
| `examples/chart-request.json` | Supported chart authoring with named-range-backed series and factual chart readback |
| `examples/pivot-request.json` | Range-backed pivot authoring, pivot-table inspection, and pivot-health analysis |
| `examples/package-security-inspect-request.json` | Encrypted workbook open plus factual package-security inspection |
| `examples/source-backed-input-request.json` | File-backed text, formula, and binary payload authoring with drawing-payload extraction |
| `examples/large-file-modes-request.json` | Low-memory `STREAMING_WRITE` execution plus open-time recalculation flagging |
| `examples/assertion-request.json` | Ordered mutate-then-verify flow with first-class assertion steps and verbose journaling |
| `examples/workbook-health-request.json` | Compact no-save health workflow combining sheet summary, formula health, aggregate workbook findings, and cell readback |

Run any JSON fixture with:

```bash
./gradlew :cli:run --args="--request examples/<file>.json"
```

Examples that persist a workbook write to `generated-workbooks/` beside the request file; for the
checked-in repository fixtures that means `examples/generated-workbooks/`. The no-save examples
return JSON only.

The Java example is compiled and executed by `:authoring-java:test` and demonstrates how the
`authoring-java` module emits one canonical `WorkbookPlan` without dropping to raw JSON. The
public walkthrough for that surface lives in [JAVA_AUTHORING.md](./JAVA_AUTHORING.md).

---

## Jackson 3.x Package Notes

GridGrind uses Jackson 3.x (`tools.jackson.core:jackson-databind:3.1.2`). The package split:

| Surface | Package |
|:--------|:--------|
| Annotations (`@JsonTypeInfo`, `@JsonSubTypes`, `@JsonProperty`) | `com.fasterxml.jackson.annotation.*` |
| Core / databind APIs (`JsonMapper`, `JacksonException`) | `tools.jackson.databind.*` / `tools.jackson.core.*` |

The annotation module kept its original Java package in Jackson 3.x. Only the core/databind
implementation packages changed. Spotless will revert incorrect `tools.jackson.annotation.*`
imports back to `com.fasterxml.jackson.annotation.*` automatically.
