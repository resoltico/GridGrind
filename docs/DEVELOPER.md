---
afad: "3.3"
version: "1.0.0"
domain: DEVELOPER
updated: "2026-03-24"
route:
  keywords: [gridgrind, build, gradle, architecture, coverage, jacoco, pmd, errorprone, spotless, java26, engine, protocol, cli]
  questions: ["how do I build gridgrind", "how do I run tests", "what is the gridgrind architecture", "how are quality gates configured", "what are the coverage requirements"]
---

# Developer Reference

**Purpose**: Build, test, architecture, and quality gate reference for GridGrind contributors.
**Prerequisites**: Java 26 (auto-provisioned via Gradle toolchains), Gradle wrapper.

---

## Architecture

GridGrind is a three-module Gradle project:

```
engine/     Workbook-core abstractions on top of Apache POI.
            No JSON, no transport, no protocol concerns.

protocol/   The AI-agent-facing contract: request and response models,
            protocol versioning, structured failures, JSON encoding,
            and execution semantics.

cli/        Thin transport adapter. Reads a protocol request from stdin
            or --request file, delegates to protocol, writes the response.
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

## Build Commands

```bash
# Run all tests
./gradlew test

# Run tests + enforce coverage floors
./gradlew coverage

# Build the executable fat JAR
./gradlew :cli:shadowJar

# Run a request file
./gradlew :cli:run --args="--request examples/budget-request.json"

# Format and lint check (Spotless)
./gradlew spotlessCheck

# Apply Spotless formatting
./gradlew spotlessApply

# Run PMD static analysis
./gradlew pmdMain pmdTest

# Run all checks
./gradlew check
```

---

## Coverage Requirements

JaCoCo enforces these floors on every `./gradlew coverage` run:

| Module | Line Coverage | Branch Coverage |
|:-------|:-------------|:----------------|
| `engine` | 100% | 95% |
| `protocol` | 100% | 95% |
| `cli` | 100% | 95% |

---

## Quality Gates

All gates run as part of `./gradlew check`.

### Error Prone

Compile-time static analysis. These checks are promoted to errors (build fails):

- `BadImport`
- `BoxedPrimitiveConstructor`
- `CheckReturnValue`
- `EqualsIncompatibleType`
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

Google Java Format 1.35.0. Removes unused imports. Enforced on every commit via `spotlessCheck`.
Run `./gradlew spotlessApply` to auto-format before committing.

### JaCoCo

100% line coverage, 95% branch coverage across all modules. The coverage report is generated at
`build/reports/jacoco/` in each module after `./gradlew coverage`.

---

## Execution Semantics

Operations run first. Analysis runs second. Persistence happens last.

If analysis fails, the request returns a structured error and no workbook is written. This gives
agents deterministic failure semantics: a failure never leaves a partially-written file behind.

---

## Workflow Fixtures

These runnable examples cover the core operation surface:

| File | What It Tests |
|:-----|:-------------|
| `examples/budget-request.json` | Range write, style, formula, analysis in one request |
| `examples/live-workflow-create.json` | Multi-sheet workbook with cross-sheet formulas and aggregations |
| `examples/live-workflow-revise.json` | Reopen, revise, recalculate, reanalyze |

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
