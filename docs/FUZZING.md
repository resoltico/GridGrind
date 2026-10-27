# Fuzzing

Fuzzing belongs to the root Gradle build and uses Java 26 with the pinned Jazzer driver.
The generators and workbook invariants live in src/fuzz/java; committed raw inputs live in
src/fuzz/resources/seeds. No separate build, launcher scripts, promotion metadata or run
history is required.

Run deterministic regression replay with:

```bash
./gradlew fuzzReplay
```

The same replay runs as part of ./gradlew check. It discovers every committed input for all
four harnesses and fails if any harness has no inputs. Ordinary JUnit runs also exercise the
generators and invariant checks. Replay does not launch the native fuzz driver.

Active fuzzing is an explicit operation:

```bash
./gradlew fuzz -Pharness=protocol-request -PfuzzSeconds=60
```

Available harnesses are protocol-request, protocol-workflow, engine-command-sequence and
xlsx-roundtrip. The default active duration is 60 seconds. The standalone driver mutates raw
bytes through FuzzTarget; active runs and deterministic replay use the same Java input cursor
and invariant methods. This implements the documented S2 fallback, rather than depending on
Jazzer's JUnit native bootstrap.

Working corpora are under build/fuzz/corpus/<harness>; finding artifacts are under
build/fuzz/findings/<harness>. Preserve a useful finding there before cleaning build outputs.
To retain a regression, copy its exact input into the matching committed seeds directory and
run fuzzReplay. Do not commit generated workbooks or active-run corpora.

Expected malformed-request and invalid-sequence classifications are normal outcomes.
Unexpected exceptions and invariant violations fail the replay or active run. Generated
workflow grants authorize the workflow's exact operations, owned source and publication;
authority failures cannot count as successful workflow exercise.

Full active campaigns and mutation campaigns are separate from ordinary checks. During the
foundations redesign they run only after all implementation is complete, as requested.
