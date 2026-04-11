# Jazzer

Local-only Jazzer fuzzing layer for GridGrind.

This nested build is intentionally separate from the main product build. It is not part of root
Gradle `check`, root coverage, or GitHub CI. Root [check.sh](../check.sh) does invoke the nested
Jazzer `check` workflow sequentially as the supported whole-repo local gate.

The nested build imports the root version catalog and shared included build logic, so dependency
versions and Gradle task conventions stay aligned with the main repository while Jazzer remains an
independent composite build. The committed
`src/main/resources/dev/erst/gridgrind/jazzer/support/jazzer-topology.json` file is the single
source of truth for harness keys, class names, task names, and working directories.

Use the root Gradle wrapper, not Brew `gradle`, and make sure the active shell resolves to Java 26
before running any Jazzer command. See [docs/DEVELOPER_JAVA.md](../docs/DEVELOPER_JAVA.md).

## Supported Scripts

- `jazzer/bin/regression`
- `jazzer/bin/fuzz-protocol-request`
- `jazzer/bin/fuzz-protocol-workflow`
- `jazzer/bin/fuzz-engine-command-sequence`
- `jazzer/bin/fuzz-xlsx-roundtrip`
- `jazzer/bin/fuzz-all`
- `jazzer/bin/status`
- `jazzer/bin/report`
- `jazzer/bin/list-findings`
- `jazzer/bin/list-corpus`
- `jazzer/bin/replay`
- `jazzer/bin/promote`
- `jazzer/bin/refresh-promoted-metadata`
- `jazzer/bin/clean-local-findings`
- `jazzer/bin/clean-local-corpus`

Use the scripts for active fuzzing and day-to-day Jazzer operator work. Raw Gradle remains useful
only for deterministic nested-build verification (`test`, `check`). The scripts provide:
- the local run lock
- per-target run history
- latest-summary artifacts
- path normalization for replay and promotion
- the active-fuzz duration watchdog
- active-fuzz `--no-daemon` launch isolation plus interrupt and timeout cleanup
- explicit separation between generated local corpus and committed custom seeds
- replay of the public defaulted-field contract for promoted JSON examples, including omitted
  optional fields that must still parse successfully, plus top-level `formulaEnvironment`
  payloads that must preserve their request shape
- replay of the advanced factual readback contract for promoted JSON examples, including workbook
  protection, rich comments, advanced print setup, structured style colors or gradients,
  autofilter criteria or sort state, and table metadata
- replay of the advanced workbook-core mutation contract for promoted JSON examples, including
  password-bearing protection, formula-defined named ranges, targeted formula evaluation, and
  explicit formula-cache clearing

Active fuzz launcher tasks now preload a tiny project-owned premain agent before `JazzerHarnessRunner`
starts. That bridge publishes JVM startup instrumentation to Byte Buddy up front so Java 26 active
fuzzing does not depend on a late external self-attach.

Active fuzzing is local-only. GitHub Actions must stay on deterministic verification only, and the
active harness runner now rejects live fuzzing when `GITHUB_ACTIONS=true`.

## Common Commands

Run the deterministic Jazzer support tests:

```bash
./gradlew --project-dir jazzer test --console=plain
```

Run the nested Jazzer verification baseline:

```bash
./gradlew --project-dir jazzer check --console=plain
```

This runs the deterministic Jazzer support tests plus regression replay of the committed seed
floor. Regression replay is isolated per harness under the hood rather than running all four
Jazzer harnesses through one Gradle `Test` task, and promoted-input semantic replay now lives
only in those dedicated regression runner tasks rather than inside `PromotionMetadataTest`.
Regression replay and `jazzer/bin/replay` use a project-owned pure-Java scalar replay cursor
instead of Jazzer's native `withJavaData(...)` path, so replay stays stable in a fresh JVM even
when upstream native replay loading is temperamental.
It does not start active fuzzing. It now emits `[JAZZER-PULSE]` lines while support tests and
regression harnesses advance so whole-repo `./check.sh` runs can show live Stage 2 progress
instead of going quiet.

Run one harness:

```bash
jazzer/bin/fuzz-protocol-workflow -PjazzerMaxDuration=30m --console=plain
```

This is the supported live-fuzz operator surface. The wrapper records run history and summaries
around the underlying Gradle task, forces active fuzz onto `--no-daemon`, and tears down the
Gradle client process tree on interrupt or timeout so a supported local run does not leave a live
harness behind.

Run all four active harnesses sequentially:

```bash
jazzer/bin/fuzz-all -PjazzerMaxDuration=5m --console=plain
```

Inspect the latest state:

```bash
jazzer/bin/status --console=plain
jazzer/bin/report protocol-workflow --console=plain
jazzer/bin/list-corpus protocol-workflow --console=plain
```

`list-corpus` shows two different inventories:
- generated local corpus under `jazzer/.local/`
- committed custom seeds under `jazzer/src/fuzz/resources/.../*Inputs/...`

Replay and promote:

```bash
jazzer/bin/replay protocol-workflow <input-path> --console=plain
jazzer/bin/promote protocol-workflow <input-path> set_cell_failure_case --console=plain
jazzer/bin/refresh-promoted-metadata --console=plain
```

Replay the committed seed floor:

```bash
jazzer/bin/regression --console=plain
```

`jazzer/bin/regression` keeps a single public operator command, but the nested build now executes
the four harnesses as isolated regression launcher tasks before writing the aggregate `regression`
summary.

## Progress Pulses

Long-running Jazzer verification surfaces operator-readable progress with stable pulse prefixes:

- `[JAZZER-PULSE] support-tests ...` for deterministic support-test start, per-test completion,
  and top-level class completion
- `[JAZZER-PULSE] regression-target ...` for each committed-seed replay harness
- `[JAZZER-PULSE] regression-input ...` for each committed promoted input replayed by a
  regression harness
- `[JAZZER-PULSE] harness-class ...` for standalone active-fuzz execution of a single harness

Root `./check.sh` also emits `[CHECK-PULSE]` lines every 15 seconds while a stage is active and
`[CHECK-DIAG]` lines if a stage stops making progress long enough to trigger automatic stall
diagnostics. A diagnosed stalled stage is then terminated with a `[CHECK-STALL]` line instead of
waiting indefinitely.

Inspect active findings versus expected-invalid and replay-clean artifacts:

```bash
jazzer/bin/status --console=plain
jazzer/bin/report xlsx-roundtrip --console=plain
```

`status` and `report` count only unexpected failures as active findings. Older local crash files
that now replay as expected-invalid or replay cleanly are shown separately.

Clean local state:

```bash
jazzer/bin/clean-local-findings
jazzer/bin/clean-local-corpus
```

## Tuning

Active fuzz scripts support:

- `-PjazzerMaxDuration=<duration>`
- `-PjazzerMaxExecutions=<count>`

If neither is set, Jazzer's default `@FuzzTest` duration applies.

## Expected Warning Noise

Two warning families are currently expected from upstream tooling during local Jazzer runs:

- `sun.misc.Unsafe` deprecation warnings from Jazzer `0.30.0` on Java 26
- `JUnit arguments ... can not be serialized as fuzzing inputs. Skipped.` during active fuzzing of
  `FuzzedDataProvider` harnesses

These warnings do not indicate a GridGrind workbook bug and do not affect replay, promotion, or
latest-summary correctness.

## Storage

Generated local state stays under:

- `jazzer/.local/run-lock/`
- `jazzer/.local/runs/*/.cifuzz-corpus/`
- `jazzer/.local/runs/*/latest.log`
- `jazzer/.local/runs/*/latest-summary.json`
- `jazzer/.local/runs/*/latest-summary.txt`
- `jazzer/.local/runs/*/telemetry/`
- `jazzer/.local/runs/*/history/`
- `jazzer/.local/runs/*/findings/`

Nothing under `jazzer/.local/` should be committed.

Committed regression inputs and promotion metadata live under:

- `jazzer/src/fuzz/resources/.../*Inputs/...`
- `jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata/`

See [docs/DEVELOPER_JAZZER_COVERAGE.md](../docs/DEVELOPER_JAZZER_COVERAGE.md) for the current
committed seed inventory and per-harness counts.

Run root Gradle verification and nested Jazzer verification sequentially, not in parallel.
They share the same workspace and composite-build outputs, and parallel runs can create stale or
misleading local build state.

## Deterministic Gradle Verification

Raw Gradle remains available only for deterministic nested-build verification:

```bash
./gradlew --project-dir jazzer test --console=plain
./gradlew --project-dir jazzer check --console=plain
```

For active fuzzing, use `jazzer/bin/*` and nothing else. That is the only supported operator path.

## Full References

- [docs/DEVELOPER_JAZZER.md](../docs/DEVELOPER_JAZZER.md)
- [docs/DEVELOPER_JAZZER_OPERATIONS.md](../docs/DEVELOPER_JAZZER_OPERATIONS.md)
- [docs/DEVELOPER_JAZZER_COVERAGE.md](../docs/DEVELOPER_JAZZER_COVERAGE.md)
