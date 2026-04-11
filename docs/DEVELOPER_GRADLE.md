---
afad: "3.5"
version: "0.37.0"
domain: DEVELOPER_GRADLE
updated: "2026-04-11"
route:
  keywords: [gridgrind, gradle, build-logic, composite-build, version-catalog, jazzer, buildsrc, toolchain, configuration-cache, verification]
  questions: ["how is the gridgrind gradle build structured", "why does gridgrind use gradle/build-logic instead of buildSrc", "how does the nested jazzer build consume the root project", "where are shared gradle conventions defined", "what should we review in the gradle setup"]
---

# Gradle Setup Reference

**Purpose**: Explain how GridGrind's Gradle system is arranged, why it is arranged that way, and
what contributors should review before changing it.
**Companion references**: [DEVELOPER.md](./DEVELOPER.md),
[DEVELOPER_JAVA.md](./DEVELOPER_JAVA.md), [DEVELOPER_JAZZER.md](./DEVELOPER_JAZZER.md)

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
protocol/
cli/
jazzer/
├── settings.gradle.kts
└── build.gradle.kts
```

Each layer owns a different concern:

- root product build: builds and verifies `engine`, `protocol`, and `cli`
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
local `engine` and `protocol` modules without publishing snapshots. This keeps Jazzer iteration
fast and ensures fuzzing runs against the exact working tree under review.

### One dependency authority

The root version catalog in `gradle/libs.versions.toml` is the shared dependency authority. The
nested Jazzer build imports that catalog instead of repeating overlapping coordinates locally. That
avoids silent version skew between the main product modules and Jazzer support code.

### Thin module build scripts

Large `.gradle.kts` files are hard to test, hard to refactor, and easy to let drift into mixed
configuration-plus-implementation blobs. GridGrind therefore keeps reusable typed logic in
`gradle/build-logic` and keeps consumer scripts thin. `jazzer/build.gradle.kts` is intentionally a
single plugin application for exactly that reason.

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
- do not run root and nested Jazzer Gradle builds in parallel against the same workspace

---

## Stable Invariants

These are the Gradle-level invariants worth preserving:

- `engine`, `protocol`, and `cli` remain ordinary root subprojects
- `jazzer/` remains a nested build, not a root subproject
- `gradle/build-logic` remains the only home for shared typed Gradle logic
- the nested Jazzer build imports `../gradle/libs.versions.toml`
- shared pulse scheduling lives in one base implementation, with build-specific listeners layered on
  top
- root product-module quality policy lives in convention plugins, not in a root `subprojects {}`
  block
- Jazzer-specific PMD and JaCoCo scopes live beside Jazzer task registration in the shared build
  logic
- root `./gradlew check` stays focused on the product modules
- active Jazzer fuzzing stays local-only; GitHub Actions must not become a live-fuzz surface
- root `./check.sh` remains the supported whole-repo gate that sequences root verification, Jazzer
  verification, packaging, and Docker smoke checks

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

If Jazzer topology changed, also run at least one live `jazzer/bin/*` command so the operator path
is exercised in the same shape contributors will actually use.
