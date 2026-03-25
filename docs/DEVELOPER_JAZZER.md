---
afad: "3.4"
version: "0.5.0"
domain: DEVELOPER_JAZZER
updated: "2026-03-25"
route:
  keywords: [gridgrind, jazzer, fuzz, fuzzing, developer, local-only, regression, corpus, crash, inputs, composite-build, gradle, junit, xlsx, architecture]
  questions: ["how does jazzer fit into gridgrind", "where does jazzer live in this repo", "how should gridgrind wire jazzer", "where do jazzer corpus files go", "is jazzer part of ci", "how do we separate fuzzing from the normal test suite", "what is the jazzer architecture for gridgrind", "what are the jazzer policies in this project"]
---

# Jazzer Developer Reference

**Purpose**: Authoritative architecture, policy, storage, and operating reference for the GridGrind
Jazzer layer.
**Prerequisites**: [DEVELOPER.md](./DEVELOPER.md) for the main build, module architecture, and
quality gates.

---

## Status and Authority

This file is the canonical Jazzer design document for GridGrind.

As of `2026-03-25`:

- Jazzer is **approved as a future local-only fuzzing layer**.
- Jazzer is **not yet implemented** in the repository.
- No `jazzer/` nested build exists yet.
- No Jazzer tasks, harnesses, corpora, or helper code exist yet.

Everything in this file is the **target contract** for how Jazzer must be introduced. Until the
corresponding files exist, treat this document as architecture and policy, not as an inventory of
implemented behavior.

If future implementation diverges from this file, update the implementation or update this file.
They must not drift.

---

## Decision Summary

GridGrind will use Jazzer as a **separate, local-only, explicitly invoked fuzzing layer**.

The approved shape is:

- a standalone nested Gradle build rooted at `jazzer/`
- invoked via the repository's existing root Gradle wrapper
- isolated from the root `check`, `coverage`, and `./check.sh` flows
- absent from GitHub CI and GitHub artifact workflows
- focused on `.xlsx` behavior only

The main deterministic JUnit suite remains the project's primary correctness contract. Jazzer adds
deeper local exploration; it does not replace ordinary tests.

---

## Goals

The Jazzer layer exists to:

- explore malformed and semi-valid request inputs more broadly than hand-written tests can
- explore ordered operation sequences that are expensive to cover manually
- explore `.xlsx` save/reopen round-trip edge cases
- discover unexpected combinations and structural edge conditions
- turn meaningful findings into durable regressions in the normal test suite

---

## Non-Goals

The Jazzer layer does **not** exist to:

- replace unit tests
- replace integration tests
- become part of the root build's correctness gate
- become part of GitHub CI
- publish artifacts anywhere
- fuzz non-`.xlsx` formats
- fuzz macros, `.xlsm`, `.xls`, encryption, or signing workflows

The `.xlsx` scope lock applies to Jazzer exactly as it applies to the rest of GridGrind.

---

## Why This Architecture

### Why Jazzer

Jazzer is the chosen fuzzing engine because it:

- integrates naturally with Gradle and the JUnit Platform
- supports both regression replay and active fuzzing through the same harness model
- is actively maintained
- fits both raw request fuzzing and structured sequence fuzzing

### Why Jazzer Must Be Separate

Jazzer must not be added as an ordinary part of the main build because the root build has very
different goals:

- deterministic execution
- strict 100% coverage gates
- fast feedback on normal development work
- CI suitability

Active fuzzing has the opposite characteristics:

- opt-in execution
- long runtimes
- local state generation
- exploratory rather than gate-oriented behavior

Keeping the fuzz layer separate is not a convenience choice. It is a correctness and maintenance
choice.

### Why `jazzer/` Must Be a Nested Build

A nested build is the cleanest way to preserve separation.

It gives us:

- physical separation in the tree
- independent task lifecycle
- no accidental inclusion in root `check`
- no accidental coupling to root JaCoCo rules
- direct access to the local checkout through Gradle composite build substitution

### Why Fuzz Harnesses Must Not Live Under `src/test`

Fuzz harnesses are not ordinary tests. They differ in:

- invocation style
- runtime profile
- seed and corpus handling
- crash artifact handling
- ownership of local state

Putting them under product-module `src/test` would blur the line between:

- public, deterministic product regressions
- local exploratory fuzz infrastructure

GridGrind must keep those concerns visually and operationally separate.

### Why Generated Corpora and Crash Files Must Stay Local

Generated Jazzer state is local machine state, not source code.

It is:

- high-churn
- machine-specific
- often noisy
- usually temporary

That state should not enter Git history. Only intentionally promoted regressions or curated seeds
belong in versioned source.

---

## Compatibility Baseline

The Jazzer layer should align with the same core runtime baseline as the main project.

| Component | Version | Notes |
|:----------|:--------|:------|
| Java | 26 | Match the repository toolchain baseline. |
| Gradle | 9.4.1 | Match the repository wrapper. |
| JUnit Jupiter | 6.0.3 | Match the repository test stack. |
| JUnit Platform Launcher | 6.0.3 runtime line | Match the resolved JUnit 6 runtime line. |
| Jazzer | 0.30.0 | Approved target baseline. |

Research findings captured for future implementation:

- Jazzer's documentation still frames JUnit support in "JUnit 5" terms.
- Jazzer's published `jazzer-junit` artifact still declares older transitive JUnit dependencies.
- A local throwaway smoke project verified that Jazzer `0.30.0` runs successfully with:
  - Gradle `9.4.1`
  - JUnit `6.0.3`
- Gradle conflict resolution upgraded Jazzer's older transitive JUnit dependencies to the JUnit 6
  line declared by the build.

Implication:

- the Jazzer build must pin JUnit 6 explicitly
- the Jazzer build must not rely on Jazzer's transitive JUnit versions

---

## Operating Modes

Jazzer must be used in two clearly separated modes.

### Regression Mode

Regression mode is the default, deterministic mode used to replay curated seeds and promoted
reproducers.

Characteristics:

- suitable for a local `check` inside the nested Jazzer build
- should finish in predictable time
- should not generate exploratory corpus churn
- should not require long-running local sessions

Typical use:

- confirm that known findings stay fixed
- confirm that harnesses still compile and execute
- validate code-defined seeds and curated fixtures

### Active Fuzzing Mode

Active fuzzing mode is the exploratory mode enabled by `JAZZER_FUZZ=1`.

Characteristics:

- local only
- opt-in
- potentially long-running
- generates local corpus and finding artifacts
- not suitable for root `check`, root `coverage`, or GitHub CI

Typical use:

- explore new edge cases
- search for unexpected operation interactions
- grow local corpora while investigating a concern

### Mode Boundary Rule

Regression mode is the only Jazzer mode that may be part of the nested Jazzer build's own local
`check` task. Active fuzzing mode must always remain an explicitly invoked local action.

---

## Repository Topology

### Approved Root Placement

The Jazzer layer belongs at:

```text
/Users/erst/Tools/GridGrind/jazzer/
```

This keeps it:

- close to the code it fuzzes
- clearly outside `engine`, `protocol`, and `cli`
- easy to omit from the main build lifecycle

### Root-Level Files That May Reference Jazzer

Only a small set of root-level files should know about Jazzer:

| Path | Role |
|:-----|:-----|
| [docs/DEVELOPER_JAZZER.md](/Users/erst/Tools/GridGrind/docs/DEVELOPER_JAZZER.md) | Canonical Jazzer design and operating reference. |
| [.gitignore](/Users/erst/Tools/GridGrind/.gitignore) | Must ignore local Jazzer state such as `/jazzer/.local/`. |
| [.gitattributes](/Users/erst/Tools/GridGrind/.gitattributes) | Already treats `.xlsx` as binary; extend only if future committed Jazzer fixtures need additional binary handling. |

The following root files must remain operationally independent of Jazzer:

| Path | Constraint |
|:-----|:-----------|
| [settings.gradle.kts](/Users/erst/Tools/GridGrind/settings.gradle.kts) | Must not include `jazzer`. |
| [build.gradle.kts](/Users/erst/Tools/GridGrind/build.gradle.kts) | Must not depend on Jazzer tasks. |
| [check.sh](/Users/erst/Tools/GridGrind/check.sh) | Must not invoke Jazzer tasks. |
| GitHub workflows under `.github/workflows/` | Must not invoke active Jazzer tasks. |

---

## Exact Planned Tree

This is the approved target tree for the Jazzer layer and its related local state.

```text
GridGrind/
├── engine/
├── protocol/
├── cli/
├── docs/
│   └── DEVELOPER_JAZZER.md
├── jazzer/
│   ├── README.md
│   ├── settings.gradle.kts
│   ├── build.gradle.kts
│   ├── gradle.properties
│   ├── bin/
│   │   ├── regression
│   │   ├── fuzz-protocol-request
│   │   ├── fuzz-protocol-workflow
│   │   ├── fuzz-engine-command-sequence
│   │   ├── fuzz-xlsx-roundtrip
│   │   ├── fuzz-all
│   │   ├── clean-local-corpus
│   │   └── clean-local-findings
│   ├── src/
│   │   └── fuzz/
│   │       ├── java/
│   │       │   └── dev/erst/gridgrind/jazzer/
│   │       │       ├── support/
│   │       │       │   ├── FuzzDataDecoders.java
│   │       │       │   ├── OperationSequenceModel.java
│   │       │       │   ├── WorkbookInvariantAssertions.java
│   │       │       │   └── XlsxRoundTripAssertions.java
│   │       │       ├── protocol/
│   │       │       │   ├── ProtocolRequestFuzzTest.java
│   │       │       │   └── OperationWorkflowFuzzTest.java
│   │       │       └── engine/
│   │       │           ├── WorkbookCommandSequenceFuzzTest.java
│   │       │           └── XlsxRoundTripFuzzTest.java
│   │       └── resources/
│   │           └── dev/erst/gridgrind/jazzer/
│   │               └── fixtures/
│   │                   ├── json/
│   │                   └── xlsx/
│   └── .local/
│       └── runs/
│           ├── protocol-request/
│           │   ├── .cifuzz-corpus/
│           │   └── ...
│           ├── protocol-workflow/
│           │   ├── .cifuzz-corpus/
│           │   └── ...
│           ├── engine-command-sequence/
│           │   ├── .cifuzz-corpus/
│           │   └── ...
│           └── xlsx-roundtrip/
│               ├── .cifuzz-corpus/
│               └── ...
└── ...
```

### Tree Rules

- `jazzer/` is the only top-level directory for Jazzer code and Jazzer-specific local state.
- `jazzer/src/fuzz/` is the only approved committed source root for Jazzer harness code.
- `jazzer/src/fuzz/resources/.../fixtures/` is the only approved committed fixture location for
  Jazzer.
- `jazzer/.local/` is the only approved location for generated local fuzz state.
- `jazzer/src/test/` is intentionally unused.
- `jazzer/src/main/` is intentionally unused.

---

## Build Architecture

### Core Rule

`jazzer/` is its own Gradle build.

It must not be included in the root [settings.gradle.kts](/Users/erst/Tools/GridGrind/settings.gradle.kts).

### `jazzer/settings.gradle.kts`

`jazzer/settings.gradle.kts` should:

- declare its own build name
- configure the Java toolchain resolver
- include the main GridGrind build as a composite build

Conceptually:

```kotlin
plugins {
    id("org.gradle.toolchains.foojay-resolver-convention") version "1.0.0"
}

rootProject.name = "GridGrind-Jazzer"
includeBuild("..")
```

### Why Composite Build Substitution Is Required

The Jazzer build should depend on the local GridGrind modules through Gradle composite build
substitution, not through published artifacts.

That means the Jazzer build should be able to declare dependencies such as:

```kotlin
dependencies {
    add("fuzzImplementation", "dev.erst.gridgrind:protocol")
    add("fuzzImplementation", "dev.erst.gridgrind:engine")
}
```

and let `includeBuild("..")` substitute those dependencies with the local checkout.

This is required because it gives us:

- zero publish step
- no stale `mavenLocal()` artifacts
- no machine-global state dependency
- direct fuzzing against the live working tree

### The Jazzer Build Must Not Reuse Root `buildSrc`

The Jazzer build must not apply the main repository's `gridgrind.java-conventions` plugin from
[buildSrc/src/main/kotlin/gridgrind.java-conventions.gradle.kts](/Users/erst/Tools/GridGrind/buildSrc/src/main/kotlin/gridgrind.java-conventions.gradle.kts).

Reasons:

- that convention plugin is tuned for the product build and its coverage gates
- it increases the risk of accidental coupling to root `check` semantics
- the Jazzer layer needs a smaller, local-only lifecycle model

If the Jazzer build needs helper conventions, they should be defined inside `jazzer/build.gradle.kts`
or inside a Jazzer-local included build in the future.

### Wrapper Policy

The root Gradle wrapper remains the only wrapper in the repository.

Use:

```bash
./gradlew --project-dir jazzer <task>
```

Do not add a second wrapper under `jazzer/`.

Reasons:

- one Gradle version to manage
- one wrapper security surface
- less repository noise

---

## Dependency Model

The Jazzer build must pin its own dependencies explicitly.

### Required Direct Dependencies

| Dependency | Purpose |
|:-----------|:--------|
| `org.junit:junit-bom:6.0.3` | JUnit 6 alignment. |
| `org.junit.jupiter:junit-jupiter` | Jupiter API and engine. |
| `org.junit.platform:junit-platform-launcher` | Explicit launcher presence. |
| `com.code-intelligence:jazzer-junit:0.30.0` | Jazzer JUnit integration. |
| `dev.erst.gridgrind:protocol` | Request-level fuzzing surface. |
| `dev.erst.gridgrind:engine` | Direct engine-level fuzzing surface. |

`dev.erst.gridgrind:cli` is optional and should be added only if CLI-specific fuzzing is later
approved.

### Dependency Rules

- Pin JUnit 6 explicitly.
- Do not trust Jazzer's transitive JUnit declarations as the chosen version source.
- Do not depend on `mavenLocal()`.
- Do not depend on published GridGrind artifacts.
- Do not introduce extra libraries until a specific harness need exists.
- Put Jazzer harness dependencies on the `fuzz` source set configurations such as
  `fuzzImplementation` and `fuzzRuntimeOnly`, not on unrelated source set configurations.

---

## Source Set and Package Layout

### Source Set

The committed Jazzer harness source set is named `fuzz`.

Approved committed roots:

- `jazzer/src/fuzz/java`
- `jazzer/src/fuzz/resources`

### Forbidden Roots

Do not place Jazzer harness code in:

- `jazzer/src/test/java`
- `jazzer/src/main/java`
- `engine/src/test/java`
- `protocol/src/test/java`
- `cli/src/test/java`

The only exception is a finding that has been deliberately promoted into the main deterministic
suite. Once promoted, it is no longer Jazzer-only knowledge.

### Package Layout

Approved package families:

- `dev.erst.gridgrind.jazzer.support`
- `dev.erst.gridgrind.jazzer.protocol`
- `dev.erst.gridgrind.jazzer.engine`

Purpose by family:

| Package | Role |
|:--------|:-----|
| `support` | Decoders, models, shared invariants, shared assertions. |
| `protocol` | Request-level and workflow-level fuzz harnesses. |
| `engine` | Engine-level sequence and `.xlsx` round-trip harnesses. |

---

## Task Architecture

### Philosophy

The Jazzer build should expose:

- a deterministic local regression layer
- several explicit active fuzzing tasks
- explicit cleanup tasks for local state

### Approved Task Set

| Task | Purpose | Fuzzing mode |
|:-----|:--------|:-------------|
| `check` | Jazzer-local sanity gate: compile harnesses and run regression-mode replay only. | No |
| `jazzerRegression` | Run all Jazzer harnesses in regression mode only. | No |
| `fuzzProtocolRequest` | Active fuzzing of request parsing and request validation. | Yes |
| `fuzzProtocolWorkflow` | Active fuzzing of ordered protocol workflows. | Yes |
| `fuzzEngineCommandSequence` | Active fuzzing of engine command sequences. | Yes |
| `fuzzXlsxRoundTrip` | Active fuzzing of `.xlsx` create/mutate/save/reopen invariants. | Yes |
| `fuzzAllLocal` | Convenience umbrella for all active local fuzz tasks. | Yes |
| `cleanLocalCorpus` | Delete generated corpora under `jazzer/.local/`. | No |
| `cleanLocalFindings` | Delete local crash files and other local finding state under `jazzer/.local/`. | No |

### Execution Rules

- Active fuzzing tasks must never be dependencies of root `check`.
- Active fuzzing tasks must never be dependencies of root `coverage`.
- Active fuzzing tasks must never be invoked by [check.sh](/Users/erst/Tools/GridGrind/check.sh).
- Active fuzzing tasks must never be wired into GitHub CI.
- Each active fuzzing task must target one harness family only.
- Each active fuzzing task must run with `maxParallelForks = 1`.
- Each active fuzzing task must set `JAZZER_FUZZ=1`.
- Each active fuzzing task must use a task-specific working directory under `jazzer/.local/runs/`.
- Regression-mode tasks must not set `JAZZER_FUZZ=1`.
- The working directory choice is part of the architecture, not an incidental implementation
  detail.

### Why One Harness Family per Task

Jazzer fuzzing mode executes one fuzz target per run. One-task-per-family keeps:

- run state separated
- corpora separated
- failures easier to triage
- local operating commands predictable

### Approved Working Directory Mapping

| Task | Working Directory |
|:-----|:------------------|
| `fuzzProtocolRequest` | `jazzer/.local/runs/protocol-request/` |
| `fuzzProtocolWorkflow` | `jazzer/.local/runs/protocol-workflow/` |
| `fuzzEngineCommandSequence` | `jazzer/.local/runs/engine-command-sequence/` |
| `fuzzXlsxRoundTrip` | `jazzer/.local/runs/xlsx-roundtrip/` |

This mapping matters because Jazzer writes generated corpus under `.cifuzz-corpus/` relative to the
working directory. By pinning each task to its own working directory, GridGrind keeps generated
corpora and local findings isolated by concern.

---

## Storage Model

### Core Principle

Generated fuzzing state is local operational data, not source code.

It belongs under:

```text
jazzer/.local/
```

### Approved Storage Layout

| Path | Contents | Versioned |
|:-----|:---------|:----------|
| `jazzer/.local/runs/protocol-request/` | Working directory for protocol request fuzzing. | No |
| `jazzer/.local/runs/protocol-workflow/` | Working directory for workflow fuzzing. | No |
| `jazzer/.local/runs/engine-command-sequence/` | Working directory for engine sequence fuzzing. | No |
| `jazzer/.local/runs/xlsx-roundtrip/` | Working directory for `.xlsx` round-trip fuzzing. | No |
| `jazzer/.local/runs/*/.cifuzz-corpus/` | Generated coverage corpora. | No |
| `jazzer/.local/runs/*/` | Crash files and other task-local outputs. | No |
| `jazzer/build/` | Nested-build compilation, reports, and test-result outputs. | No |

### Git Policy

The root [.gitignore](/Users/erst/Tools/GridGrind/.gitignore) should ignore:

```text
/jazzer/.local/
```

No generated Jazzer corpus, crash file, or task-local finding artifact should be committed.

### Spill Prevention Rules

The Jazzer layer must not allow generated state to spill into the general repository tree.

That means:

- no root-level `.cifuzz-corpus/`
- no generated crash files under `engine/`, `protocol/`, or `cli/`
- no live fuzz output under `docs/` or `examples/`
- no default use of committed `*Inputs` directories as active write targets

The task working directory architecture is what prevents this spill.

---

## Seeds, Fixtures, Corpora, and Findings

### Preferred Seed Strategy

Prefer code-defined seeds through JUnit parameter sources such as `@MethodSource`.

Reasons:

- reproducible
- reviewable
- source-controlled
- less filesystem coupling
- avoids accidental spill of live fuzz findings into versioned directories

### Fixture Strategy

Committed deterministic fixtures may live under:

```text
jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/fixtures/
```

Approved fixture types:

- small JSON examples
- small `.xlsx` examples
- curated edge-case workbooks

### Default Rule for `*Inputs` Directories

Do **not** create committed Jazzer `*Inputs` directories by default.

Reason:

- Jazzer writes discovered failing inputs into its inputs directory when one exists
- a committed inputs directory makes it easy for local crash artifacts to spill into the versioned
  tree

Therefore the default policy is:

- no committed `*Inputs` directories
- no versioned live crash-output directories
- seeds come from code or from neutral fixture locations
- local findings stay under `jazzer/.local/`

### Committed vs Local-Only Inventory

| Category | Committed | Local Only |
|:---------|:----------|:-----------|
| Harness source code | Yes | No |
| Shared decoder/assertion support code | Yes | No |
| Curated JSON fixtures | Yes | No |
| Curated `.xlsx` fixtures | Yes | No |
| Generated `.cifuzz-corpus/` contents | No | Yes |
| Crash files from active fuzzing | No | Yes |
| Temporary reproducers during triage | No by default | Yes |
| Promoted deterministic regressions in main suite | Yes | No |

### Promotion Policy for Findings

A meaningful Jazzer finding should not remain only as a local crash file.

Preferred promotion order:

1. add a deterministic regression test in the main suite
2. add a curated code-defined Jazzer seed if it still strengthens fuzz exploration
3. add a reviewed committed fixture only when a file fixture is the right long-term form

The main deterministic suite remains the authoritative home for product regressions.

---

## Harness Families

### Protocol Request Fuzzing

Location:

```text
jazzer/src/fuzz/java/dev/erst/gridgrind/jazzer/protocol/ProtocolRequestFuzzTest.java
```

Purpose:

- fuzz raw and semi-structured request inputs
- harden JSON decoding
- harden request validation
- ensure malformed inputs become structured failures rather than uncontrolled crashes

Primary assertions:

- no VM crash
- no uncaught internal exception leakage
- invalid requests fail within allowed problem families
- `.xlsx`-only policy remains enforced

### Protocol Workflow Fuzzing

Location:

```text
jazzer/src/fuzz/java/dev/erst/gridgrind/jazzer/protocol/OperationWorkflowFuzzTest.java
```

Purpose:

- fuzz ordered request operation sequences
- explore valid and invalid combinations
- verify stop semantics, persistence semantics, and error classification boundaries

Primary assertions:

- ordered workflows either succeed or fail in an allowed way
- invalid combinations fail in documented categories
- analysis failures do not silently persist output
- operation order matters only where the contract says it matters

### Engine Command Sequence Fuzzing

Location:

```text
jazzer/src/fuzz/java/dev/erst/gridgrind/jazzer/engine/WorkbookCommandSequenceFuzzTest.java
```

Purpose:

- fuzz direct engine command sequences below the protocol layer
- exercise state transitions closer to the workbook engine

Primary assertions:

- engine invariants hold after successful sequences
- invalid sequences fail in the expected exception families
- no state corruption escapes subsequent verification

### `.xlsx` Round-Trip Fuzzing

Location:

```text
jazzer/src/fuzz/java/dev/erst/gridgrind/jazzer/engine/XlsxRoundTripFuzzTest.java
```

Purpose:

- fuzz workbook creation, mutation, save, reopen, and structural verification

Primary assertions:

- expected-valid workbooks reopen successfully
- saved workbooks preserve internal structural invariants
- merged regions, sheet order, dimensions, and freeze panes remain coherent

---

## Support Code

Support code belongs under:

```text
jazzer/src/fuzz/java/dev/erst/gridgrind/jazzer/support/
```

Approved support responsibilities:

| Type | Responsibility |
|:-----|:---------------|
| `FuzzDataDecoders` | Decode `FuzzedDataProvider` data into constrained domain inputs. |
| `OperationSequenceModel` | Build structured valid and invalid operation sequences. |
| `WorkbookInvariantAssertions` | Shared workbook-state assertions. |
| `XlsxRoundTripAssertions` | Shared save/reopen assertions. |

### Decoder Policy

For GridGrind, prefer:

- `FuzzedDataProvider`
- explicit decoders
- bounded sequence models
- deliberate generation of valid and invalid state transitions

Do not rely heavily on reflective arbitrary-object generation for workbook workflows. GridGrind's
interesting fuzz surface is structured and ordered; explicit modeling is easier to reproduce and
evolve.

---

## Contracts

### Contract 1: The Main Suite Remains Primary

The main JUnit suite remains the source of truth for:

- exact public behavior
- exact error-code assertions
- CI regressions
- product contract coverage

Jazzer augments that suite. It does not replace it.

### Contract 2: Jazzer Is Local Only

Jazzer is intentionally:

- local
- opt-in
- developer-invoked
- absent from GitHub CI

### Contract 3: The Root Build Must Stay Clean

The root build must remain operationally independent of Jazzer.

That means:

- root `settings.gradle.kts` does not include `jazzer`
- root `check` does not depend on Jazzer
- root `coverage` does not depend on Jazzer
- root `check.sh` does not invoke Jazzer

### Contract 4: `.xlsx` Only

The Jazzer layer fuzzes GridGrind's approved `.xlsx` surface only.

Out of scope:

- `.xls`
- `.xlsm`
- macros
- macro preservation
- encryption
- signed workbook workflows

### Contract 5: Findings Must Be Promoted

Real findings must be turned into durable project knowledge, not left as anonymous local crash
files forever.

Preferred long-term home:

1. deterministic main-suite regression test
2. curated Jazzer seed or fixture if it still adds fuzzing value
3. local crash file only while actively investigating

---

## Procedures

### Procedure: Add a New Harness

1. Choose the layer: `protocol`, `engine`, or `.xlsx` round-trip.
2. Add the harness under `jazzer/src/fuzz/java/...`.
3. Add or update support code under `support/` only if shared logic is warranted.
4. Add one dedicated active fuzzing task for that harness family.
5. Assign the harness a dedicated `jazzer/.local/runs/<task>/` working directory.
6. Add or update a convenience script under `jazzer/bin/` if the task deserves a stable entrypoint.
7. Update this document so the architecture record stays in sync.

### Procedure: Run Regression Mode

Use:

```bash
./gradlew --project-dir jazzer jazzerRegression
```

Intent:

- replay curated seeds
- replay promoted deterministic reproductions
- confirm harnesses still build and execute

### Procedure: Run Active Fuzzing

Use a dedicated task such as:

```bash
./gradlew --project-dir jazzer fuzzProtocolWorkflow
```

Rules:

- run locally only
- expect long runtimes
- keep scope narrow
- triage outputs before promoting anything

### Procedure: Triage a Finding

1. confirm reproducibility
2. reduce to the smallest understandable reproducer
3. decide whether it is:
   - a product defect
   - a harness defect
   - an invalid assumption
4. if it is a product defect, add or plan a deterministic regression in the main suite
5. keep the local crash artifact only as long as it remains useful

### Procedure: Promote a Finding

Preferred path:

1. fix the product defect
2. add a deterministic main-suite regression test
3. optionally add a Jazzer seed if it strengthens future exploration
4. update this document if the finding changes policy or architecture

### Procedure: Prune Local State

Delete only under:

```text
jazzer/.local/
```

Do not clean Jazzer state by touching unrelated parts of the repository tree.

---

## What Must Not Happen

The following are explicitly forbidden:

- adding `jazzer` as a root subproject
- placing Jazzer harnesses under `engine/src/test`, `protocol/src/test`, or `cli/src/test`
- wiring active fuzzing into root `check`
- wiring active fuzzing into root `coverage`
- invoking active Jazzer tasks from GitHub CI
- committing generated corpora
- committing local crash files
- creating committed `*Inputs` directories by default
- depending on `publishToMavenLocal` for Jazzer-to-GridGrind wiring
- adding a second Gradle wrapper under `jazzer/`

---

## Relationship to Existing Developer Docs

This file is the primary Jazzer reference.

[DEVELOPER.md](./DEVELOPER.md) remains the reference for:

- the main build
- root quality gates
- module architecture
- root developer commands

Any other document that mentions Jazzer should summarize briefly and defer to this file for the
full contract.

---

## Proposed Commands

Canonical commands:

```bash
./gradlew --project-dir jazzer check
./gradlew --project-dir jazzer jazzerRegression
./gradlew --project-dir jazzer fuzzProtocolRequest
./gradlew --project-dir jazzer fuzzProtocolWorkflow
./gradlew --project-dir jazzer fuzzEngineCommandSequence
./gradlew --project-dir jazzer fuzzXlsxRoundTrip
./gradlew --project-dir jazzer fuzzAllLocal
```

Convenience script names:

```bash
jazzer/bin/regression
jazzer/bin/fuzz-protocol-request
jazzer/bin/fuzz-protocol-workflow
jazzer/bin/fuzz-engine-command-sequence
jazzer/bin/fuzz-xlsx-roundtrip
jazzer/bin/fuzz-all
```

The scripts are convenience wrappers only. The Gradle tasks remain the canonical interface.

Wrapper script rules:

- keep scripts thin
- make scripts delegate to the canonical Gradle tasks
- do not duplicate logic in scripts that belongs in Gradle task configuration
- do not let scripts choose alternate storage locations that bypass this document
- do not let scripts invoke root `check` as part of fuzzing workflows

---

## Implementation Checklist

When Jazzer is actually introduced, the minimum implementation set is:

1. create `jazzer/` as a nested Gradle build
2. wire `includeBuild("..")` in `jazzer/settings.gradle.kts`
3. pin JUnit 6 and Jazzer dependencies in `jazzer/build.gradle.kts`
4. create the `fuzz` source set
5. add the first harness family and its dedicated task
6. add `/jazzer/.local/` to root `.gitignore`
7. keep root `settings.gradle.kts`, root `build.gradle.kts`, and [check.sh](/Users/erst/Tools/GridGrind/check.sh) operationally unchanged
8. update this file if any practical implementation detail differs from the current plan

---

## Maintenance Rules

Update this file whenever any of the following change:

- the Jazzer tree layout
- task names
- working-directory mapping
- source-set shape
- dependency alignment strategy
- seed and fixture policy
- finding-promotion policy
- CI boundary
- `.xlsx` scope policy

This file must remain ahead of or in lockstep with implementation.

---

## Quick Rules For Agents

When working on Jazzer-related tasks in GridGrind:

1. treat this file as the primary Jazzer architecture reference
2. assume Jazzer is local-only unless the user explicitly changes that policy
3. keep active fuzzing separate from root `check`, `coverage`, and CI
4. put committed harness code only under `jazzer/src/fuzz/`
5. put generated fuzz state only under `jazzer/.local/`
6. prefer code-defined seeds over committed Jazzer inputs directories
7. promote meaningful findings into deterministic main-suite tests whenever possible
8. stay inside the `.xlsx` scope lock

---

## External References

These sources informed the approved design:

- Jazzer README:
  [https://github.com/CodeIntelligenceTesting/jazzer](https://github.com/CodeIntelligenceTesting/jazzer)
- Jazzer `@FuzzTest` Javadoc:
  [https://codeintelligencetesting.github.io/jazzer-docs/jazzer-junit/com/code_intelligence/jazzer/junit/FuzzTest.html](https://codeintelligencetesting.github.io/jazzer-docs/jazzer-junit/com/code_intelligence/jazzer/junit/FuzzTest.html)
- Jazzer JUnit artifact metadata:
  [https://repo1.maven.org/maven2/com/code-intelligence/jazzer-junit/0.30.0/jazzer-junit-0.30.0.pom](https://repo1.maven.org/maven2/com/code-intelligence/jazzer-junit/0.30.0/jazzer-junit-0.30.0.pom)

These sources are inputs. GridGrind policy is defined by this document.
