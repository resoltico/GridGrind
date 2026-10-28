# Fuzzing

The root Gradle build uses JQF 3 on Java 27. The build prepares the checksum-pinned instrumentation
agent with ASM9 visitor APIs so sealed-class metadata survives instrumentation. The active agent
and runtime use that same prepared JAR, which includes upstream licenses and modification
provenance. An unexpected instrumentation source or a rejected selected class fails the run.
The JUnit-free core and agent provide coverage feedback for GridGrind core, JSON and engine
classes. The active driver and deterministic replay share the same raw byte cursor, generators and behavioral invariants
under src/fuzz/java. Committed raw inputs live in src/fuzz/resources/seeds.

## Deterministic replay

    ./gradlew fuzzReplay

The same replay runs in ./gradlew check. It discovers all 94 committed inputs across four
harnesses and fails if any harness has no inputs. Ordinary JUnit also exercises generators,
invariants and cleanup regressions. Replay does not start active coverage-guided fuzzing.
Its temporary files belong to build/tmp/fuzzReplay; actual execution resets that owned scratch.

## Active runs

    ./gradlew fuzz -Pharness=protocol-request -PfuzzSeconds=10 -PfuzzRuns=1000

Harnesses are protocol-request, protocol-workflow, engine-command-sequence and xlsx-roundtrip.
Defaults are 60 seconds, 10,000 trials and 32,768 input bytes. Set fuzzSeconds, fuzzRuns and
fuzzMaxInputBytes to select explicit budgets. An input limit smaller than a committed seed or
a trial budget too small to admit the seed corpus is rejected.

Raw seeds enter Harness.check without framing or padding. Generated inputs use bounded byte
replacement, insertion and deletion. Every unexpected exception or invariant failure is a
finding, including IOException; a declared exception is not silently accepted. Runs stop at
the first finding or budget boundary. The seed queue is processed before duration-limited
mutation; a finding or the outer deadline can stop it. Deterministic replay remains the check
that exercises the entire committed corpus.

JQF's per-input timeout is checked through instrumented callbacks. Blocked IO or native calls
may not reach those callbacks, so the Gradle task also imposes an outer process deadline of
the requested duration plus ten seconds for initialization and shutdown. Gradle configuration,
prerequisite compilation and process termination can add command wall time. The input watchdog
is operational evidence: a timeout alone does not establish a semantic defect. Interrupted
runs preserve their in-flight input on a best-effort shutdown path.

The 32 KiB limit covers the largest committed seed, 28,841 bytes, with 3,927 bytes of mutation
room. A lower limit can reduce exploration headroom; below 28,841 it excludes an existing
protocol-request seed. The 10,000-trial default is an upper effort budget, not an established
minimum or optimum. Smaller explicit budgets are suitable for smoke checks. The other three
harnesses admit their existing seeds with `-PfuzzMaxInputBytes=1024`, but this reduces reachable
input length and is not equivalent exploration. These budgets do not guarantee a coverage level.

## Disk ownership and limits

One checkout retains its latest mutable corpus under build/fuzz/corpus/<harness>, including
JQF results. Admitted runs reset the corpus from committed seeds, including when changing
harnesses. Learned inputs do not accumulate across runs; promote useful raw inputs before
starting another run. A lock rejects concurrent active runs in the same checkout. Use isolated
checkouts for parallel runs; dependency download caches can still be shared.

The task reports a conservative corpus-payload ceiling based on the seed bytes, trial count
and maximum input size. Current defaults permit at most 327,880,208 bytes (312.7 MiB) of raw
corpus payload for the largest harness; sequential harnesses do not add their ceilings together.
This is neither a required preallocation nor a total disk quota. Filesystem block/inode overhead,
statistics, finding metadata, in-flight workbooks, JVM memory and caches are separate from
that payload bound. Lowering the trial budget reduces potential corpus growth and executions.

The standalone driver disables per-trial observation logs and all-input/current-input copies.
Coverage statistics and exact finding artifacts remain available; inherited JQF observability
properties cannot enable per-trial JSON output through this driver.

Active scratch belongs to build/fuzz/tmp and is removed after success. Failed or interrupted
runs retain scratch for diagnosis. Per-input cleanup and workflow construction failures clear
their owned files. Cleanup attempts every root and suppresses secondary errors onto the
original finding, preserving its identity.

Findings belong to build/fuzz/findings/<harness>. Unreviewed files there block a later run
before existing corpora or scratch are reset. Move useful evidence out before resuming; copy
only exact raw input bytes into the matching committed seed directory for regression replay.
Do not commit generated workbooks or mutable corpora.

## Scope

The driver checks explicit GridGrind invariants with branch/call coverage feedback. It does
not provide generic security-sink sanitizers or comparison/value-profile guidance. Path,
authority, XML and secret-handling contracts also have independent deterministic tests.
Malformed-request and invalid-sequence classifications are expected outcomes; unexpected
errors remain findings. Generated workflow grants authorize exact operations, owned sources
and publication, so authority failures cannot count as successful workflow execution.

Active fuzzing and mutation reporting are opt-in and separate from ordinary deterministic checks.
