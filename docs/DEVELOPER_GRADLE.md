---
afad: "3.5"
version: "0.52.0"
domain: DEVELOPER_GRADLE
updated: "2026-04-18"
route:
  keywords: [gridgrind, gradle, build-logic, composite-build, version-catalog, jazzer, buildsrc, toolchain, configuration-cache, verification]
  questions: ["how is the gridgrind gradle build structured", "why does gridgrind use gradle/build-logic instead of buildSrc", "how does the nested jazzer build consume the root project", "where are shared gradle conventions defined", "what should we review in the gradle setup"]
---

# Gradle Setup Reference

**Purpose**: Explain how GridGrind's Gradle system is arranged after the workstation-level Java and wrapper setup from [DEVELOPER_JAVA.md](./DEVELOPER_JAVA.md) is already in place.
**Companion references**: [DEVELOPER.md](./DEVELOPER.md),
[DEVELOPER_JAVA.md](./DEVELOPER_JAVA.md), [DEVELOPER_JAZZER.md](./DEVELOPER_JAZZER.md)

---

## Canonical Execution

GridGrind's machine-level setup rule is simple:
- use `./gradlew` for every repo build command
- treat `gradle` on `PATH` as outside the supported GridGrind workflow
- let the wrapper download the official Gradle distribution pinned by the repository
- keep the repository checkout on the local Mac filesystem as part of the normal supported setup

The wrapper version is currently `9.4.1`, as declared in
[gradle/wrapper/gradle-wrapper.properties](../gradle/wrapper/gradle-wrapper.properties).

This file therefore documents build architecture and ownership boundaries, not how to install a
global Gradle command on a machine.

Wrapper integrity is part of the standard setup:
- `gradle/wrapper/gradle-wrapper.properties` pins the distribution URL and its
  `distributionSha256Sum`
- `.github/workflows/gradle-wrapper-validation.yml` validates wrapper changes in GitHub
- contributors should treat wrapper-file edits as supply-chain-sensitive changes, not as routine noise

Full verification is part of that local-filesystem rule:
- Gradle project cache and JaCoCo execution data both rely on file locking
- mounted external volumes on macOS can reject those locks with `Operation not supported`
- if that happens, move the repository to local disk instead of standardizing a cache or build-dir
  relocation workaround as the normal workflow

---

## System Map

GridGrind has three distinct Gradle layers:

1. the root product build
2. the shared included build logic
3. the nested Jazzer build

```text
settings.gradle.kts
build.gradle.kts
gradle/
├── libs.versions.toml
└── build-logic/
    ├── settings.gradle.kts
    ├── build.gradle.kts
    └── src/main/kotlin/dev/erst/gridgrind/buildlogic/
        ├── GridGrindJavaConventionsPlugin.kt
        ├── GridGrindJazzerConventionsPlugin.kt
        ├── ScheduledPulseTestListener.kt
        └── ...
engine/
contract/
executor/
authoring-java/
cli/
jazzer/
├── settings.gradle.kts
└── build.gradle.kts
```

Each layer owns a different concern:

- root product build: builds and verifies `engine`, `contract`, `executor`, `authoring-java`,
  and `cli`
- shared included build logic: houses reusable Gradle plugins and typed build logic
- nested Jazzer build: runs Jazzer support tests, regression replay, and local fuzzing flows

The root build intentionally does not include `jazzer/` as a normal subproject. Jazzer remains a
separate nested build because its runtime model, local state, and operator flows are intentionally
different from the main product modules.

---

## Why It Is Set Up This Way

### Shared included build logic instead of `buildSrc`

GridGrind used to carry Gradle helpers in root and nested `buildSrc` directories. That arrangement
had two problems:

- root and Jazzer build logic could drift independently
- deleted helper classes could survive as stale compiled artifacts in local Gradle state

The current setup replaces both `buildSrc` trees with one explicit included build under
`gradle/build-logic`. That gives the repository one home for shared plugins, one review surface
for Gradle behavior, and one place to fix infrastructure concerns such as test pulses or cleanup
tasks.

The included build also cleans its compile output directories before recompiling. That is a
deliberate defense against stale hidden classes surviving after source deletion. Kotlin
incremental compilation is intentionally disabled in `gradle/build-logic` as well: the included
build is small, and deterministic recompilation is more important there than micro-optimizing
plugin compile time.

### Composite build for Jazzer

`jazzer/settings.gradle.kts` uses `includeBuild("..")` so the nested build can consume the live
local `engine`, `contract`, and `executor` modules without publishing snapshots. This keeps Jazzer
iteration fast and ensures fuzzing runs against the exact working tree under review.

### One dependency authority

The root version catalog in `gradle/libs.versions.toml` is the shared dependency authority. The
nested Jazzer build imports that catalog instead of repeating overlapping coordinates locally. That
avoids silent version skew between the main product modules and Jazzer support code.

Jackson note:
- `tools.jackson.core:jackson-databind` 3.x intentionally still depends on
  `com.fasterxml.jackson.core:jackson-annotations`
- do not attempt to migrate GridGrind model annotations to a nonexistent
  `tools.jackson.annotation` namespace
- keep the comment beside the version-catalog entry and the JSON round-trip regressions in place
  so this upstream Jackson rule does not get "fixed" incorrectly later

### Thin module build scripts

Large `.gradle.kts` files are hard to test, hard to refactor, and easy to let drift into mixed
configuration-plus-implementation blobs. GridGrind therefore keeps reusable typed logic in
`gradle/build-logic` and keeps consumer scripts thin. `jazzer/build.gradle.kts` is intentionally a
single plugin application for exactly that reason.

### Coverage gate protocol

GridGrind's JaCoCo wiring must be stricter than JaCoCo's defaults because this repository already
contains more than one local `Test` task in at least one module.

Rules:
- never rely on an unnamed JaCoCo `limit {}` block for coverage meaning
- always set `counter = "LINE"` and `counter = "BRANCH"` explicitly for verification rules
- always set `value = "COVEREDRATIO"` explicitly for those rules
- treat omitted `counter` as invalid build logic because JaCoCo defaults that case to
  instruction coverage, and 100% instruction coverage does not prove 100% branch coverage
- wire both `jacocoTestReport` and `jacocoTestCoverageVerification` to the same execution-data
  scope so reporting and verification cannot disagree
- collect every local `build/jacoco/*.exec` file for a module instead of hardcoding only
  `build/jacoco/test.exec`
- make coverage verification depend on `tasks.withType<Test>()` so any added `Test` task must run
  before the gate evaluates
- when a canonical module is intentionally exercised by downstream consumers, include those
  downstream `build/jacoco/*.exec` files in that module's verification scope instead of pretending
  module-local tests are the whole truth
- at the root aggregated-report layer, collect every subproject `build/jacoco/*.exec` file instead
  of assuming one `test.exec` per module
- make the root `coverage` and `jacocoAggregatedReport` tasks derive their participating
  subprojects and task dependencies dynamically instead of hardcoding today's module names

Why this rule exists:
- GridGrind's `executor` module defines a dedicated `parityTest` task in addition to the normal
  `test` task
- GridGrind's `contract` module is now validated both by local contract tests and by downstream
  `executor` plus `cli` tests after the module split, so contract coverage must include that
  consumer execution data explicitly
- if JaCoCo only reads `test.exec`, parity-only execution becomes invisible to the gate
- if JaCoCo omits `limit.counter`, the gate can appear green on instruction coverage while real
  branch gaps remain hidden
- when previously unseen uncovered code appears after fixing JaCoCo wiring, treat that as the gate
  becoming truthful rather than as the code suddenly regressing

Repository-specific note:
- GridGrind must keep collecting all local `build/jacoco/*.exec` files precisely because
  multi-`Test` execution already exists here
- `executor:parityTest` is not a special exception to remember manually; the build logic must make
  it impossible to forget
- the nested Jazzer build remains intentionally separate from root product-module coverage; its own
  `./gradlew --project-dir jazzer check` is the authoritative Jazzer coverage gate

---

## Ownership Boundaries

Use this routing table before changing the build:

| If you are changing... | Change here first |
|:-----------------------|:------------------|
| root project membership, plugin resolution | `settings.gradle.kts` |
| thin root build wiring | `build.gradle.kts` |
| repository-wide project-file formatting and aggregated coverage | `gradle/build-logic/.../GridGrindRootConventionsPlugin.kt` |
| shared Java subproject conventions | `gradle/build-logic/.../GridGrindJavaConventionsPlugin.kt` |
| shared Jazzer build behavior, Jazzer task registration, Jazzer PMD and coverage profiles, cleanup tasks | `gradle/build-logic/.../GridGrindJazzerConventionsPlugin.kt` |
| dependency versions shared across product and Jazzer | `gradle/libs.versions.toml` |
| nested Jazzer plugin wiring or imported catalogs | `jazzer/settings.gradle.kts` |
| Jazzer harness and run-target topology | `jazzer/src/main/resources/dev/erst/gridgrind/jazzer/support/jazzer-topology.json` |

Rules:

- do not add reusable typed logic back into module-local `.gradle.kts` files
- do not reintroduce `buildSrc`
- do not hardcode overlapping dependency versions inside `jazzer/`
- do not make the root build depend on active fuzzing tasks
- do not wire active Jazzer fuzz tasks into GitHub Actions; GitHub must remain deterministic-only
- do not assume Bash 4+ array semantics in `jazzer/bin/*`; the supported macOS operator surface is
  stock `/bin/bash` 3.2 under `set -u`
- do not run root and nested Jazzer Gradle builds in parallel against the same workspace

---

## Stable Invariants

These are the Gradle-level invariants worth preserving:

- `engine`, `contract`, `executor`, `authoring-java`, and `cli` remain ordinary root subprojects
- `jazzer/` remains a nested build, not a root subproject
- `gradle/build-logic` remains the only home for shared typed Gradle logic
- the nested Jazzer build imports `../gradle/libs.versions.toml`
- shared pulse scheduling lives in one base implementation, with build-specific listeners layered on
  top
- root product-module quality policy lives in convention plugins, not in a root `subprojects {}`
  block
- handwritten production Java and Kotlin sources use explicit imports only; wildcard imports fail
  the root `verifyExplicitImports` gate that `check` depends on
- Jazzer-specific PMD and JaCoCo scopes live beside Jazzer task registration in the shared build
  logic
- root `./gradlew check` stays focused on the product modules
- active Jazzer fuzzing stays local-only; GitHub Actions must not become a live-fuzz surface
- root `./check.sh` remains the supported whole-repo gate that sequences root verification, Jazzer
  verification, packaging, and Docker smoke checks
- root coverage aggregation derives its participants from JaCoCo-enabled Java subprojects and all
  of their `build/jacoco/*.exec` files rather than from a fixed module list

If a proposed change breaks one of those invariants, document the reason in code comments and in
the changelog instead of letting the system drift silently.

---

## Review Checklist

Review this setup periodically, especially after Gradle, Kotlin, or Jazzer upgrades:

- Can the shared build logic move from JVM 25 bytecode output to JVM 26 output yet?
- Is any dependency version duplicated outside `gradle/libs.versions.toml`?
- Has any typed logic crept back into a leaf `.gradle.kts` script?
- Are root and nested verification scopes still cleanly separated?
- Are long-running test pulses still emitted from shared infrastructure rather than copy-pasted
  listeners?
- Is build-logic compilation still deterministic after repeated local edits, or can Kotlin
  incremental compilation safely return?
- Does the nested Jazzer build still need to stay independent from the root project graph?
- Do the `jazzer/bin/*` wrappers still work on stock macOS `/bin/bash` 3.2 when no optional
  Gradle arguments are passed?
- Are configuration-cache or composite-build constraints forcing awkward workarounds that deserve a
  redesign instead?

This file exists so those questions can be reviewed against the current system rather than against
half-remembered history.

---

## Change Workflow

When changing the Gradle system:

1. update the owning build file or plugin, not just the consuming script
2. update companion docs if the contributor workflow or architecture changed
3. run the smallest verification that proves the change, then the supported whole-repo gate when
   the change is structural

For structural Gradle changes, the normal bar is:

```bash
./gradlew check
./gradlew --project-dir jazzer check
./check.sh
```

If Jazzer topology or `jazzer/bin/*` wrapper shell logic changed, also run at least one live
`jazzer/bin/*` command plus the zero-argument cleanup scripts so the operator path is exercised in
the same shape contributors will actually use.
