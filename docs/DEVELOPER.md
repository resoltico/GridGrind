---
afad: "3.5"
version: "0.41.0"
domain: DEVELOPER
updated: "2026-04-11"
route:
  keywords: [gridgrind, build, gradle, architecture, coverage, jacoco, pmd, errorprone, spotless, java26, engine, protocol, cli]
  questions: ["how do I build gridgrind", "how do I run tests", "what is the gridgrind architecture", "how are quality gates configured", "what are the coverage requirements"]
---

# Developer Reference

**Purpose**: Build, test, architecture, and quality gate reference for GridGrind contributors.
**Prerequisites**: Java 26 active in the current shell, Gradle wrapper.
**Java setup**: [DEVELOPER_JAVA.md](./DEVELOPER_JAVA.md)

Companion references:
- [DEVELOPER_JAVA.md](./DEVELOPER_JAVA.md)
- [DEVELOPER_GRADLE.md](./DEVELOPER_GRADLE.md)
- [DEVELOPER_JAZZER.md](./DEVELOPER_JAZZER.md)
- [DEVELOPER_JAZZER_OPERATIONS.md](./DEVELOPER_JAZZER_OPERATIONS.md)
- [DEVELOPER_JAZZER_COVERAGE.md](./DEVELOPER_JAZZER_COVERAGE.md)

---

## Architecture

GridGrind is a three-module Gradle project with compiler-enforced JPMS boundaries:

```
engine/     Apache-POI-backed workbook engine. Owns mutable workbook state,
            workbook mutation rules, and factual workbook inspection.

protocol/   The external GridGrind contract plus the single execution bridge
            into engine. Public contract code is protocol-owned and split by
            responsibility: dto, operation, read, catalog, json, and exec.
            DefaultGridGrindRequestExecutor is the only main-source class
            here that is engine-aware.

cli/        Thin transport adapter. Reads a protocol request from stdin
            or --request file, delegates to protocol, writes the response,
            and exposes help/template/catalog discovery commands.
```

The CLI is not the core. The foundation is `engine` plus `protocol`. The CLI is one adapter on
top. Future adapters (HTTP, gRPC, library embedding) can be added without touching `engine` or
`protocol`.

The enforced dependency graph is:

```text
dev.erst.gridgrind.cli -> dev.erst.gridgrind.protocol -> dev.erst.gridgrind.engine
```

`cli` does not depend on `engine`, and `protocol` does not re-export `engine`. Shared Java build
conventions enable `modularity.inferModulePath`, so the `module-info.java` descriptors in all
three modules participate in normal local builds, CI, and release verification.

---

## Foundations

| Component | Version |
|:----------|:--------|
| Java | 26 |
| Apache POI | 5.5.1 |
| Jackson Databind | 3.0.3 |
| JUnit Jupiter | 6.0.3 |
| Log4j Core | 2.25.3 |

GridGrind's runtime and product-module baseline is Java 26. The only deliberate exception is the
shared included build logic under `gradle/build-logic`, which still emits JVM 25 bytecode because
Kotlin `2.3.0` does not yet target JVM 26 directly. That build logic compiles with the Java 26
toolchain and only lowers the emitted bytecode level, so the repository no longer requires a
separate Java 25 installation.

---

## Commands

`./gradlew check` remains the root-project CI gate: Spotless formatting, Error Prone, PMD, tests,
and JaCoCo coverage verification for `engine`, `protocol`, and `cli`. `./check.sh` is the local
full-stack gate: root `check`
plus `coverage`, nested Jazzer `check`, `:cli:shadowJar`, shell syntax checks for the release-
surface scripts, and a Docker smoke test that runs the image from a non-default working directory
with weird request/response/save paths while also asserting the published help/version/response
contract semantically. `check.sh` intentionally lives in the repository root as the canonical
contributor entrypoint, while `scripts/` contains helper scripts invoked by the root gate and
GitHub workflows. Both should be clean before a release-quality change is considered done. When
Stage 1 reaches long-running Gradle `Test` tasks, the shared test conventions emit
`[GRADLE-TEST-PULSE]` lines with class-start, class-complete, and throttled test-progress facts
so the local watchdog tracks semantic execution rather than mistaking quiet test output for a
hang.

Use `./gradlew`, not Brew `gradle`. GridGrind's CLI, fat JAR, release flow, and `./check.sh` all
depend on the ambient shell `java`, so `command -v java` must resolve to Java 26 and not to the
macOS `/usr/bin/java` stub. `./check.sh` now fails fast if the shell runtime is wrong.

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
./scripts/docker-smoke.sh
```

The protocol catalog is generated from the protocol record signatures. If you change request
records, nested tagged unions, or plain input records, update the catalog summaries and keep the
field-shape output authoritative: every field should still publish required/optional status and
the exact nested/plain group accepted by polymorphic inputs. The catalog build path now uses a
small internal `protocol.catalog.gather` seam for the two cases that genuinely benefit from
Stream Gatherers: ordered uniqueness and reflected field-metadata expansion. The built-in request
template remains an ordinary constant because no domain gatherer semantics are needed there. Keep
`--print-request-template` as the smallest valid machine-readable request; workflow snippets such
as the no-save `ANALYZE_WORKBOOK_FINDINGS` lint pass belong in help text, public docs, and the
example request set instead of the emitted template itself.

---

## GitHub Workflows

Release automation is split across three workflows:

- `CI` runs the root-project `./gradlew check` gate and a separate Docker smoke job that builds
  the fat JAR and verifies the Docker image from a non-default working directory.
- `Release` builds the fat JAR, publishes the GitHub Release idempotently when a `v*` tag is
  pushed, and verifies that the published release object and `gridgrind.jar` asset exist.
- `Container` first runs the same Docker smoke script against the local Dockerfile build, then
  builds and publishes the multi-arch GHCR image, verifies with an isolated anonymous Docker
  config that both the exact version tag and `latest` are publicly pullable and runnable, and
  prunes older container package versions.

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

| Module | Line Coverage | Branch Coverage |
|:-------|:-------------|:----------------|
| `engine` | 100% | 100% |
| `protocol` | 100% | 100% |
| `cli` | 100% | 100% |

`./gradlew check` enforces the thresholds but does not write report files. Run
`./gradlew coverage` to enforce thresholds and generate HTML/XML reports at
`build/reports/jacoco/` in each module plus an aggregated cross-module report at
`build/reports/jacoco/aggregated/`.

The table above applies only to the root product modules. The nested Jazzer build has a separate
coverage contract because its local-only generator, telemetry, and operator classes are exercised
primarily through regression replay and live fuzzing rather than ordinary unit tests. Run
`./gradlew --project-dir jazzer check` to enforce Jazzer's dedicated coverage scope together with
its shared Spotless and PMD gates.

---

## Execution Semantics

Operations run first. Reads run second. Persistence happens last.

If a read fails, the request returns a structured error and no workbook is written. This gives
agents deterministic failure semantics: a failure never leaves a partially-written file behind.

---

## Workflow Fixtures

These runnable examples cover the core operation surface:

| File | What It Tests |
|:-----|:-------------|
| `examples/budget-request.json` | Range write, style, formula, workbook summary, cells, window, and schema reads |
| `examples/advanced-readback-request.json` | Rich factual readback for workbook protection, rich comments, advanced print setup, structured style colors, autofilter criteria or sort state, and advanced table metadata |
| `examples/advanced-mutation-request.json` | Workbook-core mutation parity example for password-bearing protection, formula-defined names, advanced table/autofilter mutation, advanced conditional formatting, and rich comments |
| `examples/data-validation-request.json` | Validation authoring, partial clearing, factual validation reads, and validation-health analysis |
| `examples/excel-authoring-essentials-request.json` | Hyperlink, comment, named-range authoring plus explicit metadata and named-range reads |
| `examples/file-hyperlink-health-request.json` | File-hyperlink authoring plus explicit hyperlink metadata and hyperlink-health analysis |
| `examples/formatting-depth-request.json` | Font, fill, and border styling with explicit post-mutation cell and window reads |
| `examples/introspection-analysis-request.json` | Read-heavy workbook showcasing factual reads plus finding-bearing analysis operations together |
| `examples/live-workflow-create.json` | Multi-sheet workbook with cross-sheet formulas and aggregations |
| `examples/live-workflow-revise.json` | Reopen, revise, recalculate, and reread |
| `examples/sheet-management-request.json` | Sheet copy, active/selected state, visibility, protection, and workbook-summary reads |
| `examples/structural-layout-request.json` | Merge, size, pane, zoom, print-layout, and repeating-title shaping with layout reads |

Run any fixture with:

```bash
./gradlew :cli:run --args="--request examples/<file>.json"
```

Output workbooks are written to `cli/build/generated-workbooks/`.

---

## Jackson 3.x Package Notes

GridGrind uses Jackson 3.x (`tools.jackson.core:jackson-databind:3.0.3`). The package split:

| Surface | Package |
|:--------|:--------|
| Annotations (`@JsonTypeInfo`, `@JsonSubTypes`, `@JsonProperty`) | `com.fasterxml.jackson.annotation.*` |
| Core / databind APIs (`JsonMapper`, `JacksonException`) | `tools.jackson.databind.*` / `tools.jackson.core.*` |

The annotation module kept its original Java package in Jackson 3.x. Only the core/databind
implementation packages changed. Spotless will revert incorrect `tools.jackson.annotation.*`
imports back to `com.fasterxml.jackson.annotation.*` automatically.
