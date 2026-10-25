---
afad: "4.0"
domain: DEVELOPER
updated: "2026-08-31"
route:
  keywords: [gridgrind, build, gradle, architecture, archunit, coverage, jacoco, pmd, errorprone, spotless, java26, devcontainer, zulu26, engine, contract, executor, authoring-java, cli]
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
- [ARCHITECTURE.md](./ARCHITECTURE.md)
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

GridGrind currently has four Gradle projects with compiler-enforced JPMS boundaries:

- `core`: canonical request/result types, Excel vocabulary, catalog metadata, Java authoring
  and wire codecs; it has no POI or engine dependency.
- `engine`: POI-backed execution, admission, input binding, preflight, staged publication and
  evidence. Runtime and parity tests live with the engine.
- `cli`: input/output transport, command dispatch, rendering and workspace materialization.
- `architecture-tests`: verification-only rules and focused negative controls.

See [ARCHITECTURE.md](./ARCHITECTURE.md). Product module descriptors own readability and exports;
ArchUnit checks the boundaries that JPMS cannot express. Java authors use the core-owned fluent
API and model without pulling in POI. Execution crosses the exported
`dev.erst.gridgrind.engine.api` seam with an explicit grant.

Within `engine`, the workbook runtime no longer treats `dev.erst.gridgrind.excel` as one flat
catch-all namespace. High-churn workbook subdomains now sit in dedicated `event`, `stream`,
`pivot`, `customxml`, `ooxml`, `drawing`, and `validation` subpackages so package names carry the
same theory boundaries as the code they own.

The highest-churn architecture seams are intentionally split too:
- `GridGrindProtocolCatalog` owns the top-level catalog assembly, `CatalogFieldMetadataSupport` derives nested/plain field-shape metadata from the canonical descriptor groups, and `GridGrindProtocolCatalogLookupSupport` owns exact lookup plus grouped operation-centered search behavior.
- Jazzer request generation is no longer one monolith: the `OperationSequence*` family now splits orchestration, selector helpers, mutation/inspection/command factories, observation-side assertions, bounded payload generation, and chart payload generation into focused seams, while `WorkbookInvariantChecks` is split across workbook/cell/engine-surface invariant helpers.
- `architecture-tests:architectureTest` executes the rules that JPMS cannot express. JPMS owns module readability and exports; ArchUnit protects intra-engine dependencies, API signatures (including POI and generic types), tooling gateways, closed model shapes, and admission/semantics/publication boundaries. Gradle supplies the actual compiled product roots. Focused negative controls prove the retained rules detect their defects, and the task rejects an empty discovery result. Runtime and XSSF parity tests execute in `engine`.

## Foundations

| Component | Version |
|:----------|:--------|
| Contributor devcontainer | pinned glibc base image plus Azul Zulu 26 JDK; see [DEVELOPER_DEVCONTAINER.md](./DEVELOPER_DEVCONTAINER.md) |
| Java | 26 |
| Docker runtime | Docker Desktop daemon plus `docker buildx` reachable through the active shell `docker` command; smoke and release verification use an anonymous `DOCKER_CONFIG` while still targeting the active local Docker engine |
| Architecture tests | ArchUnit 1.5.0 with its JUnit 6 integration in `architecture-tests:architectureTest` |
| Apache POI | 5.5.1 |
| Jackson Databind | 3.2.2 |
| JUnit Jupiter | 6.1.3 |
| Log4j Core | 2.26.1 |

GridGrind's runtime, product modules, and shared included build logic under `gradle/build-logic`
all target Java 26 now. The included build is no longer a JVM 25 exception: it compiles with
Kotlin `2.4.20`, emits JVM 26 bytecode directly, and stays aligned with the repository's
single Java baseline instead of carrying a separate bytecode-level footnote.
Build-task implementations and convention plugins are Kotlin; their regression tests remain
Java. The included build reads its Java target from the root `gradle.properties` pin.

Jackson dependency note: Jackson 3.x databind intentionally still uses the
`com.fasterxml.jackson.core:jackson-annotations` artifact and package namespace. That is upstream
Jackson design, not a GridGrind version-skew bug.

---

## Commands

`./gradlew check` remains the root-project code-quality gate: Spotless formatting,
explicit-import verification, executable ArchUnit rules, Error Prone with NullAway on `@NullMarked` production packages, PMD, tests, and
JaCoCo HTML/XML reports for `core`, `engine`, `architecture-tests`, and `cli`.
`./check.sh` is the supported whole-repo deterministic gate: root `check` plus `coverage`,
nested Jazzer `check`,
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
hang. The ordered stage inventory and shell-regression list are canonically owned
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
`docker buildx build --load` instead of Docker's deprecated classic builder path.
The committed contributor devcontainer is validated separately through
`./scripts/validate-devcontainer.sh`, which builds the contributor image, checks its pinned Java
and tooling surface, guards the portable devcontainer contract, and verifies the committed VS Code
overlay contract in `.devcontainer/devcontainer.json`.

```bash
# Run the local full-stack gate
./check.sh

# Generate advisory mutation reports
./gradlew mutationReport

# Run the root-project CI gate
./gradlew check

# Targeted iteration during development
./gradlew test                   # tests only, no quality gates
./gradlew coverage               # tests + advisory HTML/XML coverage reports
./gradlew architectureCheck      # mandatory ArchUnit product architecture rules
./gradlew mutationReport         # advisory PIT reports for contract and engine
./gradlew spotlessCheck          # formatting check only
./gradlew spotlessApply          # auto-format in place (run before spotlessCheck)
./gradlew pmdMain pmdTest        # PMD only

# Build artifacts
./gradlew :cli:shadowJar
./gradlew :cli:run --args="--print-grant-template --response grant.json"
# Edit grant.json to authorize examples/budget-request.json's exact resources, operations, and output.
./gradlew :cli:run --args="--request examples/budget-request.json --grant grant.json"
./gradlew :cli:run --args="--version"
./gradlew :cli:run --args="--print-request-template"
./gradlew :cli:run --args="--print-protocol-catalog"
./gradlew :cli:run --args="--print-recipe --lookup BUDGET"
./gradlew :cli:run --args="--print-recipe-catalog"
./gradlew :cli:run --args="--print-recipe-catalog --lookup DASHBOARD"
./gradlew :cli:run --args="--print-recipe --lookup DASHBOARD"
./gradlew :cli:run --args='--print-recipe-keyword-match --query "monthly sales dashboard with charts"'
./scripts/verify-cli-discovery-execution.sh ./cli/build/libs/gridgrind.jar
./scripts/docker-smoke.sh
./scripts/validate-devcontainer.sh
```

The protocol catalog is generated from the contract record signatures. If you change request
records, nested tagged unions, or plain input records, keep the field-shape output authoritative:
every field should still publish required/optional status and the exact nested/plain group
accepted by polymorphic inputs. Contract-bearing discovery prose such as low-memory mode limits,
formula parse boundaries, and workbook-health summaries now live in the canonical contract layer
instead of thin downstream string copies. `GridGrindContractText` owns stable wording and
`GridGrindExecutionModeMetadata` owns the structured EVENT_READ and STREAMING_WRITE contract
rules plus their validation messages. CLI-specific presentation lives downstream in `cli`:
`CliSurface` and `GridGrindCliHelp` own the help section labels, key/value entries, flags,
docs links, and example routing that the transport renders. `cli` owns the high-level task
descriptor DTOs plus typed discovery profiles, while `GridGrindCliRecipeRegistry`,
`GridGrindRecipeDefinition`, `GridGrindRecipeDefinitions`, the concrete
`GridGrindExampleRecipeDefinition` and `GridGrindTaskRecipeDefinition` variants, and the
per-task `*TaskRecipe` classes,
`GridGrindAdHocTaskRequestScaffolds`, and
`GridGrindRecipeKeywordMatcher` owns the downstream recipe/discovery policy: one shared published
recipe source feeds the built-in example view plus the executable task-starter view, official
task ids resolve only through that registry-owned recipe set, the ad hoc scaffolder builds only
generic unpublished task requests, and English keyword query ranking is anchored first in typed
goal/artifact metadata rather than loose notes or pitfalls. Request linting is also part of that
authoritative public surface now: `engine` exports a narrow request-doctor API alongside the
request executor, and the packaged-artifact verifiers exercise the emitted doctor report plus
every published built-in example and task starter instead of relying only on module-local tests.
The build now also includes
contract-side and CLI-side public-surface
linters that fail if docs, generated help, catalog summaries, shipped examples, or shared runtime
diagnostics mention a canonical mutation, assertion, or inspection id that is not registered in
the catalog vocabulary.
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
  black-box verifies the published CLI help, protocol catalog, task catalog, recipe,
  task keyword matcher, and doctor surfaces from both tags before pruning older container package
  versions.
- `Mutation testing` runs weekly or on manual dispatch and retains advisory PIT reports.
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

The CLI tests include a golden behaviour corpus for all shipped examples and recipes,
discovery commands, doctor problem cases, and execution failures. Each case runs in process
and through the packaged shadow JAR. Saved workbook facts are reopened through GridGrind
inspection rather than compared as binary files.

Run the corpus with `./gradlew :cli:test --tests '*GoldenCorpusTest' --no-parallel`.
To record an intentional change, add `-Pgolden.update=true`, review the diff under
`cli/src/test/resources/golden/snapshots/`, and explain each changed observation in the commit
message. Only the in-process launcher records; the packaged launcher still has to match it.
The suite includes a one-character cell-value negative control that must fail comparison.

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

- `NullAway` on `@NullMarked` production packages, using JSpecify nullness
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

Standard PMD rules complement compiler analysis and behavioural tests.

- `gradle/pmd/ruleset.xml` enforces `errorprone`, `bestpractices`, `multithreading`,
  `performance`, `security`, documentation, and the production-safe subset of `design`.
- `gradle/pmd/test-ruleset.xml` keeps test-code hygiene while allowing dense fixtures.
- Error Prone and NullAway enforce compiler-level correctness; Spotless owns formatting.
- ArchUnit checks closed domain shapes and requires concrete wire variants to be records.
  The current sealed branch hierarchy is traversed to its leaves. The rule has a negative
  control demonstrating that an enum wire alternative is rejected.
- `VerifyRuntimeLegalInventoryTask` guards the resolved runtime attribution inventory;
  `PruneJarOutputsTask` keeps packaging outputs consistent with declared artifacts.

### Spotless

Google Java Format 1.36.1. Removes unused imports. Run `./gradlew spotlessApply` to auto-format
before committing; `./gradlew spotlessCheck` (run by `check`) will fail if formatting is off.
Project-file formatting intentionally excludes local-only instruction and scratch areas,
so personal workspace state cannot destabilize the canonical quality gates.

### JaCoCo

Coverage is reported as HTML and XML. `./gradlew check` writes per-module reports;
`./gradlew coverage` also writes the aggregate under `build/reports/jacoco/aggregated/`.
Reports collect execution data from every local `Test` task, including architecture and parity
suites, and from downstream consumers where those consumers exercise the owning model or engine.

Core reports include downstream engine and CLI execution.
`engine` reports include its runtime/parity suites and downstream CLI execution data.
The root aggregate derives participants from the JaCoCo-enabled Java subprojects.

The nested Jazzer build also produces advisory coverage reports. Its `check` runs deterministic
support tests, committed-seed replay, Spotless and PMD. Reports include its full main source set;
coverage percentages do not determine success.

PIT reports are generated by `./gradlew mutationReport` for the explicit scopes in the `core`
and `engine` build files. The weekly/manual mutation workflow retains HTML/XML reports for review.
Survivors and coverage scores are advice for improving behaviour tests, with no score threshold.

Supported `jazzer/bin/*` wrappers must remain compatible with stock macOS `/bin/bash` 3.2 under
`set -u` until their replacement in the root fuzz build.

---

## Execution Semantics

Mutation steps run first in authored order, assertion and inspection steps run wherever they are
placed in authored order, and persistence happens only after every step succeeds.

If any step or calculation phase fails, the request returns a structured error and workbook
persistence is skipped. This gives agents deterministic failure semantics at the `.xlsx` output
boundary: GridGrind does not emit a partially written workbook file.

---

## Workflow Fixtures

The shipped JSON fixtures are generated from the CLI-owned example registry in
checkout-rooted form. Refresh them with
[`scripts/sync-generated-examples.sh`](../scripts/sync-generated-examples.sh), or print any one of
the artifact-native built-in examples directly with `gridgrind --print-recipe --lookup <self-contained-id>`
or materialize an asset-backed workspace with `gridgrind --materialize-recipe --lookup <id> --workspace <new-directory>`. The full
map, path-rooting rules, and verification loop live in [EXAMPLES.md](./EXAMPLES.md). Do not
hand-edit the checked-in `examples/*.json` fixtures or the generated package-security workbook;
regenerate them from the CLI-owned recipe registry so the committed examples, built-in example surface,
and verification loop stay on one source of truth. These
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

Run any JSON fixture through an explicit least-authority grant:

```bash
./gradlew :cli:run --args="--print-grant-template --response grant.json"
# Edit grant.json to authorize examples/<file>.json's exact resources, operations, and output.
./gradlew :cli:run --args="--request examples/<file>.json --grant grant.json"
```

Examples that persist a workbook write to `generated-workbooks/` beside the request file; for the
checked-in repository fixtures that means `examples/generated-workbooks/`. The no-save examples
return JSON only.

The Java example is compiled and executed by `:core:test` and demonstrates how the
core authoring API emits one canonical `WorkbookPlan` without dropping to raw JSON. The
public walkthrough for that surface lives in [JAVA_AUTHORING.md](./JAVA_AUTHORING.md).

---

## Jackson 3.x Package Notes

GridGrind uses Jackson 3.x (`tools.jackson.core:jackson-databind:3.2.0`). The package split:

| Surface | Package |
|:--------|:--------|
| Annotations (`@JsonTypeInfo`, `@JsonSubTypes`, `@JsonProperty`) | `com.fasterxml.jackson.annotation.*` |
| Core / databind APIs (`JsonMapper`, `JacksonException`) | `tools.jackson.databind.*` / `tools.jackson.core.*` |

The annotation module kept its original Java package in Jackson 3.x. Only the core/databind
implementation packages changed. Spotless will revert incorrect `tools.jackson.annotation.*`
imports back to `com.fasterxml.jackson.annotation.*` automatically.
