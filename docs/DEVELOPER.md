---
afad: "3.4"
version: "0.22.0"
domain: DEVELOPER
updated: "2026-03-31"
route:
  keywords: [gridgrind, build, gradle, architecture, coverage, jacoco, pmd, errorprone, spotless, java26, engine, protocol, cli]
  questions: ["how do I build gridgrind", "how do I run tests", "what is the gridgrind architecture", "how are quality gates configured", "what are the coverage requirements"]
---

# Developer Reference

**Purpose**: Build, test, architecture, and quality gate reference for GridGrind contributors.
**Prerequisites**: Java 26 (auto-provisioned via Gradle toolchains), Gradle wrapper.

Local-only Jazzer architecture, operations, and coverage inventories live in:
- [DEVELOPER_JAZZER.md](./DEVELOPER_JAZZER.md)
- [DEVELOPER_JAZZER_OPERATIONS.md](./DEVELOPER_JAZZER_OPERATIONS.md)
- [DEVELOPER_JAZZER_COVERAGE.md](./DEVELOPER_JAZZER_COVERAGE.md)

---

## Architecture

GridGrind is a three-module Gradle project:

```
engine/     Workbook-core abstractions on top of Apache POI.
            No JSON, no transport, no protocol concerns.

protocol/   The external GridGrind contract and execution layer:
            request and response models, the request-executor port,
            the default executor, structured failures, JSON encoding,
            protocol discovery metadata, and execution semantics.

cli/        Thin transport adapter. Reads a protocol request from stdin
            or --request file, delegates to protocol, writes the response,
            and exposes help/template/catalog discovery commands.
```

The CLI is not the core. The foundation is `engine` plus `protocol`. The CLI is one adapter on
top. Future adapters (HTTP, gRPC, library embedding) can be added without touching `engine` or
`protocol`.

---

## Foundations

| Component | Version |
|:----------|:--------|
| Java | 26 |
| Apache POI | 5.5.1 |
| Jackson Databind | 3.0.3 |
| JUnit Jupiter | 6.0.3 |
| Log4j Core | 2.25.3 |

---

## Commands

`./gradlew check` remains the root-project CI gate: Spotless formatting, Error Prone, PMD, tests,
and JaCoCo coverage verification. `./check.sh` is the local full-stack gate: root `check`
plus `coverage`, nested Jazzer `check`, `:cli:shadowJar`, shell syntax checks for the release-
surface scripts, and a Docker smoke test that runs the image from a non-default working directory
with weird request/response/save paths. `check.sh` intentionally lives in the repository root as
the canonical contributor entrypoint, while `scripts/` contains helper scripts invoked by the root
gate and GitHub workflows. Both should be clean before a release-quality change is considered done.

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
the exact nested/plain group accepted by polymorphic inputs.

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
| `examples/data-validation-request.json` | Validation authoring, partial clearing, factual validation reads, and validation-health analysis |
| `examples/excel-authoring-essentials-request.json` | Hyperlink, comment, named-range authoring plus explicit metadata and named-range reads |
| `examples/file-hyperlink-health-request.json` | File-hyperlink authoring plus explicit hyperlink metadata and hyperlink-health analysis |
| `examples/formatting-depth-request.json` | Font, fill, and border styling with explicit post-mutation cell and window reads |
| `examples/introspection-analysis-request.json` | Read-heavy workbook showcasing factual reads plus finding-bearing analysis operations together |
| `examples/live-workflow-create.json` | Multi-sheet workbook with cross-sheet formulas and aggregations |
| `examples/live-workflow-revise.json` | Reopen, revise, recalculate, and reread |
| `examples/sheet-management-request.json` | Sheet rename, delete, and reorder semantics |
| `examples/structural-layout-request.json` | Merge, size, freeze-pane shaping, and layout reads |

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
