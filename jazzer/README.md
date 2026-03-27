# Jazzer

Local-only Jazzer fuzzing layer for GridGrind.

This nested build is intentionally separate from the main product build. It is not part of root
`check`, root coverage, [check.sh](../check.sh), or GitHub CI.

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
- `jazzer/bin/clean-local-findings`
- `jazzer/bin/clean-local-corpus`

Use the scripts instead of raw Gradle when possible. They provide:
- the local run lock
- per-target run history
- latest-summary artifacts
- path normalization for replay and promotion
- explicit separation between generated local corpus and committed custom seeds

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
Jazzer harnesses through one Gradle `Test` task. It does not start active fuzzing.

Run one harness:

```bash
jazzer/bin/fuzz-protocol-workflow -PjazzerMaxDuration=30m --console=plain
```

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
```

Replay the committed seed floor:

```bash
jazzer/bin/regression --console=plain
```

`jazzer/bin/regression` keeps a single public operator command, but the nested build now executes
the four harnesses as isolated regression launcher tasks before writing the aggregate `regression`
summary.

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

The committed custom seed floor currently contains 34 inputs across the four replayable harnesses.

Run root Gradle verification and nested Jazzer verification sequentially, not in parallel.
They share the same workspace and composite-build outputs, and parallel runs can create stale or
misleading local build state.

## Raw Gradle Form

Raw Gradle still works when needed:

```bash
./gradlew --project-dir jazzer <task>
```

But raw Gradle is not the supported operator surface.

## Full References

- [docs/DEVELOPER_JAZZER.md](../docs/DEVELOPER_JAZZER.md)
- [docs/DEVELOPER_JAZZER_OPERATIONS.md](../docs/DEVELOPER_JAZZER_OPERATIONS.md)
- [docs/DEVELOPER_JAZZER_COVERAGE.md](../docs/DEVELOPER_JAZZER_COVERAGE.md)
