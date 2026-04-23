---
afad: "3.5"
version: "0.56.0"
domain: DEVELOPER_JAZZER_OPERATIONS
updated: "2026-04-17"
route:
  keywords: [gridgrind, jazzer, fuzz, operations, replay, promote, corpus, findings, summaries, telemetry]
  questions: ["how do I use the jazzer scripts", "how do I replay a jazzer input", "how do I promote a jazzer input", "where do jazzer run logs and summaries go", "how do I inspect the corpus", "how do I clean jazzer state"]
---

# Jazzer Operations Reference

**Purpose**: Day-to-day operator runbook for the GridGrind Jazzer layer.
**Architecture reference**: [DEVELOPER_JAZZER.md](./DEVELOPER_JAZZER.md)
**Coverage inventory**: [DEVELOPER_JAZZER_COVERAGE.md](./DEVELOPER_JAZZER_COVERAGE.md)

---

## Supported Scripts

| Script | Purpose |
|:-------|:--------|
| `jazzer/bin/regression` | replay all committed regression inputs |
| `jazzer/bin/fuzz-protocol-request` | fuzz raw request parsing and validation |
| `jazzer/bin/fuzz-protocol-workflow` | fuzz ordered `DefaultGridGrindRequestExecutor` workflows |
| `jazzer/bin/fuzz-engine-command-sequence` | fuzz direct workbook-command execution |
| `jazzer/bin/fuzz-xlsx-roundtrip` | fuzz `.xlsx` save and reopen invariants |
| `jazzer/bin/fuzz-all` | run all four active fuzz scripts sequentially |
| `jazzer/bin/status` | show a concise latest-summary table |
| `jazzer/bin/report` | show detailed latest-summary output |
| `jazzer/bin/list-findings` | list replayed local finding artifacts |
| `jazzer/bin/list-corpus` | list generated local corpus state and committed custom seeds |
| `jazzer/bin/replay` | replay one local input against one harness |
| `jazzer/bin/promote` | promote one input into committed regression resources |
| `jazzer/bin/refresh-promoted-metadata` | replay every promoted seed and refresh its stored replay metadata |
| `jazzer/bin/clean-local-findings` | remove local logs, summaries, telemetry, and findings |
| `jazzer/bin/clean-local-corpus` | remove local corpora |

All supported scripts use the shared lock under `jazzer/.local/run-lock/`, so only one Jazzer
command runs at a time.

---

## Choose the Surface

- `./check.sh`: supported whole-repo local gate. Runs root verification first, then nested Jazzer
  `check`, then packaging and Docker smoke.
- `./gradlew --project-dir jazzer test` and `./gradlew --project-dir jazzer check`: deterministic
  nested-build verification entrypoints. Safe for GitHub Actions because they do not start active
  fuzzing.
- `jazzer/bin/*`: the one supported Jazzer operator surface for active fuzzing, regression, replay,
  promotion, reporting, and cleanup. Active fuzz through this surface forces `--no-daemon` and
  owns interrupt and timeout teardown.
- that same wrapper surface must remain compatible with stock macOS `/bin/bash` 3.2 under
  `set -u`, including zero-argument cleanup scripts such as `jazzer/bin/clean-local-findings`

Do not run active fuzzing through raw `./gradlew --project-dir jazzer ...` tasks. Those tasks are
an implementation detail under the wrapper, not a supported fuzz operator interface.

Live fuzzing is local-only. Active harness execution now hard-fails when `GITHUB_ACTIONS=true`.

---

## Common Workflows

### Run the Full Local Gate

```bash
./check.sh --console=plain
```

This runs the root quality gates and coverage reports first, then the nested Jazzer `check`
workflow, then the CLI fat-JAR packaging step, shell syntax checks, and the Docker smoke test.
Jazzer's Stage 2 verification includes shared Spotless and PMD enforcement, deterministic support
tests, the dedicated `jazzerCoverageVerification` gate, and committed-seed regression replay. It
is the supported one-command local verification path when a change touches both the main codebase
and the Jazzer layer.

If the Docker daemon is unavailable locally, Stage 5 fails as an environment problem after the
earlier code and Jazzer stages have already completed.

### Run the Deterministic Jazzer Support Tests

```bash
./gradlew --project-dir jazzer test --console=plain
```

This runs only the deterministic nested-build test layer for Jazzer support code such as:
- summary parsing
- summary rendering
- `.xlsx` round-trip verifier behavior

It does not perform active fuzzing. During long support cases the task emits `[JAZZER-PULSE]`
class-start, test-complete, class-complete, and throttled `test-progress` lines so `./check.sh`
can distinguish active work from a real stall.

### Run One Harness

```bash
jazzer/bin/fuzz-protocol-workflow -PjazzerMaxDuration=30m --console=plain
```

What this produces:
- active fuzzing in `jazzer/.local/runs/protocol-workflow/`
- `latest.log`
- `latest-summary.json`
- `latest-summary.txt`
- `telemetry/*.json`
- a new `history/<timestamp>/` directory

For active fuzz, the wrapper also forces `--no-daemon` and tears down the launched Gradle client
tree on interrupt or timeout. That keeps the supported local operator surface from leaving a live
`JazzerHarnessRunner` behind after a canceled run.
Wrapper changes are not complete until the zero-argument maintenance scripts still run on the real
macOS shell that contributors use by default.

### Run All Active Harnesses

```bash
jazzer/bin/fuzz-all -PjazzerMaxDuration=5m --console=plain
```

`fuzz-all` is intentionally sequential. It calls the four per-harness scripts one by one so each
harness gets its own summary and history directory.

Each active fuzz launcher also preloads the project-owned `JazzerPremainAgent`, so live fuzzing on
Java 26 starts with startup-time instrumentation already published to Byte Buddy instead of
depending on a late attach.

GitHub Actions must not use this surface. If `GITHUB_ACTIONS=true`, the harness runner exits
before fuzzing starts.

### Replay Committed Regression Inputs

```bash
jazzer/bin/regression --console=plain
```

This runs all committed `*Inputs` resources in regression mode and writes the regression summary to
`jazzer/.local/runs/regression/`.

Under the hood, regression replay is isolated per harness. The public `jazzer/bin/regression`
command drives four separate regression launcher tasks and then records one aggregate regression
summary. This avoids relying on Gradle's `Test` result store for the entire committed seed floor.
Structured binary-harness replay uses GridGrind's pure-Java scalar replay cursor rather than
Jazzer's native replay bootstrap path, so replay remains stable in a fresh JVM.

`protocol-request` and `protocol-workflow` are stable Jazzer harness names for the canonical JSON
request surface. They are not module names; the live product split is `contract` plus `executor`.
Committed `protocol-workflow` binary seeds are opaque generator inputs, not human-authored semantic
scenarios. They therefore use neutral case identifiers such as `workflow_case_01.bin`, while the
authoritative decoded behavior lives in the refreshed replay metadata and replay text.

### Inspect the Current State

```bash
jazzer/bin/status --console=plain
jazzer/bin/report protocol-workflow --console=plain
jazzer/bin/list-corpus protocol-workflow --console=plain
jazzer/bin/list-findings protocol-workflow --console=plain
```

### Replay One Input

```bash
jazzer/bin/replay protocol-workflow \
  jazzer/.local/runs/protocol-workflow/.cifuzz-corpus/dev/erst/gridgrind/jazzer/protocol/OperationWorkflowFuzzTest/executeWorkflow/<sha1> \
  --console=plain
```

This does not fuzz. It replays the exact raw bytes and prints a structured summary:
- input path
- harness
- success vs expected-invalid vs unexpected-failure
- decoded operation, assertion, command, or read counts
- style kinds, source type, and persistence type where applicable
- assertion kinds where applicable
- read kinds where applicable
- response classification where applicable

Replay for `protocol-workflow`, `engine-command-sequence`, and `xlsx-roundtrip` now runs through
the project-owned deterministic replay seam. If replay disagrees with an active-fuzz finding,
treat that as a replay-contract bug to investigate immediately rather than assuming the active
finding is still current.

### Promote One Input

```bash
jazzer/bin/promote protocol-workflow \
  jazzer/.local/runs/protocol-workflow/.cifuzz-corpus/dev/erst/gridgrind/jazzer/protocol/OperationWorkflowFuzzTest/executeWorkflow/<sha1> \
  workflow_case_12 \
  --console=plain
```

Promotion writes:
- the committed regression input under `jazzer/src/fuzz/resources/.../*Inputs/...`
- promotion metadata under `jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata/...`

Promotion metadata includes:
- the replay outcome kind
- a stable replay expectation
- typed replay details
- the replay text artifact path

Promote after replay, not blindly.
For opaque binary harnesses such as `protocol-workflow`, promote with a neutral case identifier
instead of embedding assumed semantics in the file name. The replay metadata is the authoritative
description of what the current topology decodes.

### Refresh Promoted Metadata

```bash
jazzer/bin/refresh-promoted-metadata --console=plain
```

Use this after an intentional replay-shape change, such as a new command family, a new assertion
family, or an expanded
sequence-introspection model. The command replays every promoted input, rewrites the replay text,
and refreshes the stored replay outcome plus replay expectation metadata.

The refresh command is expected to rewrite many metadata files at once after a legitimate replay
contract change. Review those changes as one semantic batch, not as unrelated individual diffs.

---

## Progress Pulses

Long-running verification now emits semantic pulse lines instead of going silent:

- `jazzer check` emits `[JAZZER-PULSE] support-tests ...`, `[JAZZER-PULSE] regression-target ...`,
  and `[JAZZER-PULSE] regression-input ...`
- root `./check.sh` emits `[CHECK-PULSE] ...` with elapsed time, quiet time, stalled time, and
  the latest semantic progress marker

Interpretation rules:
- a long gap with the same latest committed input means the harness is still working on that input
  and has not yet stalled
- `quiet` means no new output; `stalled` means no new semantic progress marker
- only a `[CHECK-STALL]` line means the outer gate has decided the stage is genuinely stuck and
  has terminated it after collecting diagnostics

---

## Run Directory Layout

Each active target uses:

```text
jazzer/.local/runs/<target>/
├── .cifuzz-corpus/
├── latest.log
├── latest-summary.json
├── latest-summary.txt
├── telemetry/
│   └── <harness>.json
├── findings/
│   ├── crash-....json
│   └── crash-....txt
└── history/
    └── <timestamp>/
        ├── run.log
        ├── summary.json
        ├── summary.txt
        ├── telemetry/
        └── findings/
```

`regression/` has the same summary and telemetry model but no active corpus tree.
Its working tree also contains per-harness regression subdirectories used only during isolated
regression replay runs.

---

## What the Artifacts Mean

`latest.log`
- raw console output from the most recent run of that target

`latest-summary.json`
- machine-readable latest-summary contract
- includes run timings, exit code, corpus before/after, parsed run metrics, telemetry snapshots,
  and finding inventory

`latest-summary.txt`
- human-readable rendering of the same latest-summary
- distinguishes active findings from expected-invalid and replay-clean artifacts

`telemetry/*.json`
- per-harness semantic metrics collected during the run
- includes iteration counts, outcome counts, operation/command-kind counts, style-kind counts,
  protocol source/persistence-type counts where applicable, read-kind counts where applicable,
  error families, and protocol response families

`findings/*.json` and `findings/*.txt`
- replayed local finding artifacts derived from raw `crash-*`, `timeout-*`, `oom-*`, or `leak-*`
  files
- may include artifacts that now replay cleanly after a fix; those remain listed here until local
  cleanup removes them

`list-corpus`
- distinguishes generated local corpus from committed custom seeds explicitly
- helps explain log lines such as `seed corpus: files: ...`, which combine both sources during
  active fuzzing

`status` and `report`
- count only active findings as findings
- show expected-invalid artifacts separately when a stored raw finding now replays as a documented
  invalid-input case
- show replay-clean artifacts separately so stale local crash files are visible but not misread as
  current failures

`Recommended dictionary`
- is advisory output from libFuzzer, not a finding and not a required follow-up
- must not be copied into the repo directly from one run

Expected warning noise during local runs:
- `sun.misc.Unsafe` deprecation warnings from Jazzer `0.30.0` on Java 26
- `JUnit arguments ... can not be serialized as fuzzing inputs. Skipped.` for `FuzzedDataProvider`
  active-fuzz harnesses launched through the JUnit Platform

These warnings are upstream-tooling noise. They do not indicate a GridGrind workbook defect and
do not invalidate replay, promotion, status, or report output.

---

## Promotion Protocol

Use this exact order:

1. Replay the input with `jazzer/bin/replay`.
2. Decide whether the input is worth preserving.
3. Promote with a stable descriptive name.
4. Run `jazzer/bin/regression`.
5. If the input exposes a real GridGrind defect, convert it into a deterministic ordinary test in
   the main suite after the fix lands.

Naming rules:
- use lowercase snake case for human-authored JSON and command seeds
- use neutral `workflow_case_##` identifiers for opaque binary `protocol-workflow` seeds
- let replay metadata, not opaque binary file names, carry the authoritative decoded behavior

Good:
- `sheet_management_request`
- `workflow_case_12`
- `invalid_cell_address_case`
- `create_sheet_roundtrip_case`

Bad:
- `seed1`
- `corpus_copy`
- `set_cell_failure_case` for an opaque `protocol-workflow` binary
- `325d16aa`

Seed-quality rules:
- prefer a few semantically distinct seeds over many near-duplicates
- for `protocol-request`, prefer readable public example requests
- for the binary harnesses, prefer replay-verified promoted corpus entries
- keep at least one successful seed in every binary harness, not only invalid ones
- keep the seed floor broad enough to preserve sheet management, structural layout, formatting
  depth, authoring metadata, named-range behavior, and source/persistence-type coverage together
- treat `OperationSequenceModel` selector bytes as a stability contract for promoted binary seeds;
  if that grammar changes intentionally, refresh promotion metadata immediately and add a
  changelog-worthy note to the Jazzer developer docs

Dictionary-evaluation rules:
- only evaluate dictionary promotion after repeated similar recommendations from the same harness
- benchmark dictionary candidates against the same harness with comparable duration and settings
- reject candidate dictionaries that are large, opaque, or clearly overfit to one run
- do not promote a dictionary unless the evidence is strong enough to justify a permanent operator
  and documentation burden

---

## Cleanup Protocol

Delete local non-corpus state:

```bash
jazzer/bin/clean-local-findings
```

Delete corpora:

```bash
jazzer/bin/clean-local-corpus
```

Delete both:

```bash
jazzer/bin/clean-local-findings
jazzer/bin/clean-local-corpus
```

These commands do not touch committed regression inputs or promotion metadata.

---

## Deterministic Nested-Build Gradle Commands

Raw Gradle remains available only for deterministic nested-build verification:

```bash
./gradlew --project-dir jazzer test --console=plain
./gradlew --project-dir jazzer check --console=plain
```

For active fuzzing, use `jazzer/bin/*` and nothing else. That is the only documented operator path
that owns run locking, per-target history, latest-summary artifacts, duration control, and
interrupt cleanup.
If wrapper shell logic or topology changes, rerun at least one live `jazzer/bin/fuzz-*` command
and the zero-argument cleanup scripts before declaring the operator path verified.

---

## Warnings and Caveats

Expected:
- `sun.misc.Unsafe` deprecation warnings from Jazzer 0.30.0 on Java 26
- recommended-dictionary output after active fuzzing runs

Expected for structured replay and promotion:
- a small amount of native-access/Jazzer runtime noise may still appear

Not expected:
- missing latest-summary artifacts after a successful supported fuzz script
- missing replay output for a valid `replay` command
- missing promotion metadata after a successful `promote` command
- expected-invalid or replay-clean artifacts being counted as active findings in `status` or
  `report`
- running root Gradle verification and nested Jazzer verification in parallel without flakiness

If one of those expectations fails, treat it as a Jazzer-layer bug.

If you see `jcmd` during a Jazzer-adjacent local session, that is root `./check.sh` stall
diagnostics collecting JVM thread dumps after a monitored stage stops making semantic progress.
That diagnostic path is bounded to a small captured-process sample and now tears timed-out
diagnostic subprocess trees down authoritatively through the shared
`scripts/check-process-support.sh` helper; the supported `jazzer/bin/*` happy path does not
launch `jcmd`.

Operator discipline:
1. Prefer `./check.sh` when you want the supported sequential root-plus-nested verification flow.
2. Finish any root `./gradlew ...` or `./check.sh` run before starting a nested Jazzer build.
3. Finish any nested Jazzer run before restarting a root Gradle build.
4. Treat parallel root-plus-nested execution as unsupported local operator behavior.

If a recommended dictionary appears interesting, use this procedure:
1. Save the run log.
2. Wait for repeated recommendations from later long runs of the same harness.
3. Compare the repeated entries and keep only the clearly stable fragments.
4. Run a controlled A/B experiment for that harness with and without the candidate dictionary.
5. Only if the benefit is real and repeatable should a committed dictionary proposal be written.
