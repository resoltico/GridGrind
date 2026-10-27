# Fuzzing

The root Gradle build uses Java 26 and the pinned Jazzer driver. Generators and workbook
invariants live in src/fuzz/java; the 94 committed raw inputs live in src/fuzz/resources/seeds.

## Deterministic replay

```bash
./gradlew fuzzReplay
```

The same replay runs in ./gradlew check. It discovers every committed seed for all four
harnesses and fails if any harness has no inputs. Ordinary JUnit also exercises generators,
invariants and cleanup regressions. Replay uses the pure Java input cursor and does not
start the native driver. Its temporary files belong to build/tmp/fuzzReplay; each actual
execution resets that task-owned scratch.

## Active runs

```bash
./gradlew fuzz -Pharness=protocol-request -PfuzzSeconds=10 -PfuzzRuns=1000
```

Harnesses are protocol-request, protocol-workflow, engine-command-sequence and xlsx-roundtrip.
The standalone driver mutates bytes through FuzzTarget using the same cursor and invariants
as replay. A run stops at the first finding, the native execution budget or the time budget,
whichever comes first. Defaults are 60 seconds, 10,000 executions and 32,768 input bytes;
configure them with fuzzSeconds, fuzzRuns and fuzzMaxInputBytes respectively. A native
10-second per-input timeout reports slow or stuck cases for investigation instead of allowing
one input to block for the driver's much longer default. A timeout alone does not establish a
semantic defect. Overall time is checked between inputs, so startup and a
final in-flight input can extend wall time beyond fuzzSeconds. A smaller input
budget that excludes a committed seed is rejected, as is an execution budget too small to
admit every seed. No seed is silently skipped.

The 32 KiB input ceiling covers the largest committed seed (28,841 bytes), adding 3,927 bytes
of mutation headroom. A limit below 28,841 bytes excludes that seed; a smaller rounded budget
is possible with less exploration headroom. The 10,000-execution default roughly covers the
observed one-minute protocol-workflow run (9,290 executions), while shortening faster harnesses;
it is an exploration budget, not a coverage guarantee. Smaller explicit budgets are useful
for smoke checks. Changing either limit changes the reachable mutations or their number.

## Disk ownership and limits

One checkout retains only its latest mutable corpus under build/fuzz/corpus/<harness>.
Every admitted run resets the corpus from its committed seeds, including when changing
harnesses. Promote useful inputs before starting another run; learned corpus inputs do not
accumulate across runs. A file lock rejects concurrent active runs in one checkout. Use isolated
checkouts for parallel runs; dependency download caches can still be shared.

The maximum newly written corpus data is bounded conservatively by
(fuzzRuns + seed count + 1) × fuzzMaxInputBytes, plus the initial seed bytes. The seed-count
and one-run allowances cover native initialization and the observed iteration-boundary
overshoot. Defaults give about 314 MiB plus seed bytes for the largest 48-seed harness and
the single retained corpus; no space is preallocated. This conservative ceiling assumes every
execution adds a distinct maximum-sized input. This bounds corpus payload bytes, not allocated disk
blocks or directory metadata; filesystem overhead adds space. Observed corpora are much smaller.
Lowering fuzzRuns lowers that ceiling linearly and trades away executions; 32 KiB is already close to the
largest existing seed, so reducing execution count is the simpler way to tighten disk use.

Native temporary files, generated workbooks and extracted libraries belong to build/fuzz/tmp.
Successful runs remove this scratch; failed or killed runs retain it for diagnosis until the
next admitted run. Generator failure paths and per-input harness cleanup delete their owned
files; cleanup errors fail instead of silently accumulating leaked files. The corpus formula
is not a hard quota on in-flight workbooks, unpacked native data, JVM memory, reports or caches.
Those resources have separate lifetimes and require free filesystem space.

Finding inputs and Java reproducers go to build/fuzz/findings/<harness>. Unreviewed files in
any findings directory block another run before existing scratch or corpora are reset. Move
useful evidence out before resuming, and copy exact regression inputs into the matching
committed seed directory. Java reproducers encode inputs and can be larger than the raw-input
limit. Findings are preserved rather than pruned by age or size. Do not commit generated
workbooks, Java reproducers or mutable working corpora.

Expected malformed-request and invalid-sequence classifications are normal outcomes.
Unexpected exceptions and invariant violations fail replay or active runs. Generated workflow
grants authorize exact operations, owned sources and publication; authority failures cannot
count as successful execution. Active fuzzing and advisory mutation campaigns are opt-in and
are separate from ordinary deterministic checks.
