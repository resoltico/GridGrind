# Foundations redesign baseline

Recorded on 2026-10-25, before product changes, from
`70e3d18bb38ac66f1238a7fda28ca5b54ac68b25` on `redesign/foundations`.
The starting working tree was clean. The user's agent-guidance rewrite and
protocol-stack deletion were already committed on `main`. No root supplement
exists. `AGENTS.md` is unchanged.

## Environment

- Java and javac: Azul Zulu 26.0.1, runtime build `26.0.1+8`.
- Gradle: wrapper 9.7.1, using the committed distribution checksum.
- Docker server: 29.9.2 (reachable; image smoke not reached).

## Original command results

| Command | Exit | Wall time | Result |
| --- | --- | --- | --- |
| `./check.sh` | 1 | 213.67 s (script reports 214 s) | Hygiene passed; Stage 2 failed; Stages 3–6 not reached. |
| `./gradlew check` | 1 | 36.46 s | Same three failing tasks. |
| `./gradlew check --continue --no-parallel` | 1 | 131.22 s | Five failing tasks; other available checks exercised without exclusions. |

The original failures are:

1. `verifyControlPlaneShape`: 13 reviewed exceptions expired on 2026-09-30.
2. `verifyJavaSourceShape`: reviewed exceptions expired on 2026-09-30 and
   2026-10-15. The implementation checks the current UTC date in
   `JavaSourceShapeReviewPolicy`.
3. `build-logic:test`: all three `RepositoryPolicyFreshnessTest` methods fail
   because source-, control-plane-, and semantic-shape review dates expired.
4. `verifyJavaSemanticShape` (revealed by `--continue`): expired semantic-shape
   exceptions.
5. `contract:test` (revealed by `--continue`): two documentation audits read the
   deleted `.codex/AGENTS_EXTRA.md`. `ArchitectureSeamAuditTest` also asserts
   superseded wording in the rewritten baseline. These assertions are obsolete;
   the remaining developer-documentation and executable architecture checks stay.

PMD also emits a platform/auxClasspath warning. No PMD task failed in the
diagnostic run. No product test failure was observed outside those two obsolete
documentation audits. This does not establish that later shell, fuzz, packaging,
or container checks pass: the whole-repository gate stopped before those stages.

### Test reports from the diagnostic run

| Project / suite | Tests | Failures | Errors | Skipped |
| --- | ---: | ---: | ---: | ---: |
| build-logic / test | 25 | 3 | 0 | 0 |
| authoring-java / test | 23 | 0 | 0 | 0 |
| contract / test | 620 | 2 | 0 | 0 |
| engine / test | 1125 | 0 | 0 | 0 |
| cli / test | 455 | 0 | 0 | 0 |
| executor / architectureTest | 16 | 0 | 0 | 0 |
| executor / architectureTestContract | 1 | 0 | 0 | 0 |
| executor / parityTest | 14 | 0 | 0 | 0 |
| executor / test | 401 | 0 | 0 | 0 |
| Total | 2680 | 5 | 0 | 0 |

`excel-foundation:test` reports `NO-SOURCE`. Counts come from JUnit XML, rather
than inferring success from quiet console output.

## Authorized sequencing exception

The user explicitly approved a narrow Phase 0 exception in chat: WP-00–02
commits may retain the verified pre-existing shape-policy failures; all other
checks must pass. The golden corpus still precedes gate removal. WP-10 removes
the obsolete shape mechanisms and restores the normal green-commit requirement.
No expiration dates, clock values, task exclusions, or failure settings were
changed to manufacture a pass.

## Baseline repair verification

Removed only the hardcoded agent-guidance reads and assertions from the two
documentation audits. The user additionally confirmed that tests must not depend
on fixed `AGENTS.md` or `AGENTS_EXTRA.md` contents. A repository-wide test-source
search found no other such content dependency. Synthetic publication tests cover
archive/JAR exclusion independently of the real guidance contents.

- Focused formatting and both repaired audit classes: exit 0, Gradle reports 9 s.
- `./gradlew check --continue --no-parallel`: exit 1, 60.30 s. Only
  `verifyControlPlaneShape`, `verifyJavaSourceShape`, `verifyJavaSemanticShape`
  and the three policy-freshness tests fail. All 620 contract tests pass.
- Overall test count remains 2680, with 3 failures, 0 errors, and 0 skips.
- `git diff main -- AGENTS.md`: empty.

The repair is ready under the explicitly authorized Phase 0 exception. WP-01
(golden corpus) is next; no product refactor or gate removal has begun.

## Inventory baseline

Regenerated with the package's `tools/regenerate-inventories.sh`, outside the
repository. It matches the packaged baseline:

| Metric | Before |
| --- | ---: |
| Contract input types | 81 |
| Contract report types | 96 |
| Engine snapshot types | 51 |
| Engine definition types | 13 |
| Converter lines | 4576 |
| Files with JsonSubTypes | 99 |
| Handwritten catalog descriptor lines | 2849 |
| Contract JSON files | 104 |
| Support classes | 222 |
| Blank checks | 477 |
| Coverage-named tests | 102 |
| Shell lines | 11116 |
| Main Java lines | 128860 |
| Test Java lines | 153151 |

## Evidence retained outside the repository

`/Users/erst/Tools/gridgrind-redesign-bom/evidence/wp-00/` contains the full
command logs, original failing documentation-audit XML reports, test-count JSON,
and regenerated inventories. The unsuccessful thread-attach diagnostic ran after
the first Gradle process exited; it is not evidence of a build failure.

This temporary record is removed in WP-80 after its results are carried into
the final delivered-state report.
