# Foundations redesign spikes

Run on Java 26.0.1 on 2026-10-25. Throwaway builds lived outside the
repository in the work package's `evidence/wp-02/` directory. No runtime
library was added to the product. Full logs, JUnit XML, prototype schemas and
the exact fuzz finding are retained there; spike source and build trees are
removed after recording these results.

## S1 — Jackson metadata under JPMS: PASS

Two named modules (`spike.core`, `spike.test`) exercised a sealed union with
record variants, core-owned `Union`/`Variant`/`WireDefault` annotations, a
validating `WireString` value, and `Optional<String>`. Core used only Java types
and qualified `opens` to `tools.jackson.databind`; it did not require or import
Jackson. The JSON-side introspector and subtype registration were derived from
those annotations and `getPermittedSubclasses()`.

One JUnit test passed (1 found, 1 successful, 0 skipped/failed; 309 ms), asserting:

- union round trips, including configurable `type` and `status` discriminators;
- NON_ABSENT output and empty optionals on absent input;
- a tree default read from record-component metadata before binding;
- one generic value-string codec using the validating string constructor;
- actual reflective binding across the qualified JPMS opening;
- rejection of unknown fields and explicit null reference/primitive values.

The first prototype compile lacked JSpecify on the Jackson module path. Adding
the already-pinned JSpecify 1.0.1 corrected that setup error; the original log is
retained as `s1-attempt1.log`. The passing log is `s1.log`.

Chosen path: DEC-04's Jackson-free core. Its annotation-dependency fallback is
not needed. WP-21 and WP-22 must preserve these demonstrated behaviours.

## S2 — Jazzer JUnit: active run FAILED; documented fallback selected

A scratch Gradle root build with a `fuzz` source set used Jazzer 0.30.0 and
JUnit 6.1.3. It replayed all 48 committed protocol-request seeds plus Jazzer's
empty input through the current protocol parser, with `JAZZER_FUZZ` absent.
Regression replay passed (`s2-replay.log`, Gradle 44 s). It dynamically loaded
an agent and installed hooks; this is recorded rather than describing replay as
entirely native-free.

The opt-in active task bootstrapped libFuzzer, loaded the 48 inputs and reported
coverage, but found a pre-existing product defect after 1,524 executions / 11 s.
The Gradle run failed in 22 s instead of completing its requested 30 s:

```
IllegalArgumentException: jsonPath must not be blank
  RequestStructuralProblemSupport.copyJsonPath
  RequestInvalidJson.<init>
  RequestTolerantSyntaxParser.problem
  RequestTolerantSyntaxParser.parseObject
```

The exact input is retained outside scratch as `blank-key-syntax-finding.bin`
and in the golden doctor corpus as `problems/blank-key-syntax.bin`. The malformed
property name is accepted into a diagnostic path that its constructor rejects.
A packaged doctor invocation reproduces `INVALID_REQUEST` with unavailable JSON
location instead of preserving classified syntax findings. The existing harness
was not weakened by catching arbitrary `IllegalArgumentException`.

Chosen path: S2's documented fallback, keeping `GridGrindFuzzData` and ordinary
parameterized JUnit seed replay; opt-in active fuzzing moves to a single Gradle
`JavaExec` standalone-driver task in WP-60. A prototype fallback replay passed
all 48 seeds (Gradle 12 s); JVM library-load logging contained no Jazzer library
(`s2-fallback-replay.log`, `s2-plain-libraries.log`). No seed was skipped.

The product defect remains explicitly open for the intake work. The fallback
does not fix it or establish clean active-fuzz readiness. WP-30 must repair the
path construction and retain the regression; WP-60 must demonstrate its
standalone active task and the required clean 60-second runs per harness.

## S3 — JSON Schema: scalar/structural PASS; semantic precondition PARTIAL

A reflection prototype emitted Draft 2020-12 request and SET_CELL schemas from
the current records, sealed variants, protocol-presence/default metadata and
catalog constraints. NetworkNT 2.0.8 independently validated all 14 examples
against the request schema (535 definitions) and rejected an invalid SET_CELL
operand. Six positive/negative control pairs covered every currently permitted
`FieldConstraint` kind; neither test was skipped (`s3.log`, two tests passed,
Gradle 3 s).

| Current kind | Schema mapping |
| --- | --- |
| NonBlank | string pattern requiring a non-whitespace character |
| StringPattern | pattern |
| LengthRange | minLength / maxLength |
| NumberRange | minimum / maximum |
| Integral | integer type |
| PathSuffix | escaped, case-insensitive literal-suffix pattern |

Record properties, required fields, forbidden unknown fields, enum values,
Optional absence, metadata defaults and discriminated oneOf alternatives were
also represented. Prototype setup initially omitted BigDecimal, temporal and
short/byte scalar mappings; those generator gaps were corrected, with failing
logs retained as `s3-attempt1.log` through `s3-attempt3.log`. This is not a claim
that every current constructor invariant has already been independently mapped.
The independent validator used its default JDK regex engine; cross-runtime
ECMA-262 regex portability was not separately verified.

The sole current `OperationPrecondition`, `NO_COLUMN_EDITS_AFTER_FORMULA_AUTHORING`,
depends on ordered PRESENT/ABSENT/UNKNOWN formula state and on formulas in an
existing workbook. See `WorkbookStaticFormulaColumnValidation` and
`WorkbookStaticRequestContract.validateKnownFormulaPresence`. Document shape
alone cannot establish that external workbook state. DEC-10's narrow fallback
applies: retain this precondition in catalog metadata and static/runtime intake
validation. Do not claim the generated structural schema proves it.

### Independent validator selection

Selected for test scope in WP-30: `com.networknt:json-schema-validator:2.0.8`.
The upstream release page records the October 1 changelog / October 2 release,
more than two weeks before this spike:
[upstream release](https://github.com/networknt/json-schema-validator/releases/tag/2.0.8).
Its [versioned documentation](https://github.com/networknt/json-schema-validator/blob/2.0.8/README.md)
identifies the 2.x Jackson 2 line. The resolved scratch classpath used Jackson
2.22.1 core/databind and annotations 2.22 alongside the product's Jackson 3.2.2
and annotations 2.22. Both tests passed without linkage conflicts;
`s3-dependencies.log` records the exact resolved graph. This dependency is not
on the product runtime classpath. Pin the selected version in the catalog and
record it in package `05-BILL-OF-MATERIALS.md` when WP-30 adds it.

## Phase 0 boundary verification

For this documentation-only work package, `./gradlew check` ran in 32.12 s and
failed only at the already-approved expired source/control-plane policies and
three freshness tests. The unchanged product and golden code retain the full
WP-01 verification evidence in `GOLDEN.md`.

The required phase-end `./check.sh` ran in 47.41 s (script summary rounds to
48 s). Hygiene passed; Stage 2 failed on the same expired policies and freshness
tests. Stages 3–6 were not reached, so this run does not verify nested Jazzer,
shell/distribution checks or Docker smoke. No gate was skipped or clock changed.
The user's narrow Phase 0 exception covers these failures until WP-10 removes
the obsolete mechanisms. `AGENTS.md` remains unchanged; no supplement exists yet.

## WP-10 union-rule challenge and verification

The current mutation and inspection roots directly permit sealed subinterfaces.
A direct-record-only rule would require changing their model hierarchy before
the metadata work. After the user requested QA and proceeding by engineering
judgment, the interim rule follows permitted subclasses recursively and requires
every concrete wire alternative to be a record. Its positive control traverses a
sealed branch; one negative-control fixture proves rejection of both an enum and
an open interface branch. WP-22 must enforce the final annotation-owned shape.

The removed mechanisms comprise 14 Java implementations, nine tests of those
implementations, three policy files and the dedicated PMD profile. The CPD
dependency was used only by the deleted duplication task and is also removed.
Standard PMD, Spotless, Error Prone/NullAway, legal inventory and JAR pruning stay.

The first focused compile failed because the retained `buildSrc` rejection check
still used the repository layout variable; that variable was restored. The first
full check then failed on six PMD findings in the new fixture (205.13 s). A
focused repair passed in 9 s. The final `./gradlew check --no-parallel` passed
in 192.30 s: 2803 tests, zero failures/errors/skips, including all 141 golden
tests. The old coverage gates were still active and passed.

The current implementation and guidance contain zero references to the deleted
policy mechanisms. A literal repository-wide text search still finds release
history in `docs/CHANGELOG_ARCHIVE.md` and the temporary original baseline record.
Those are retained as historical evidence, following the baseline's instruction
to retain valid historical references. They are not active checks or guidance;
the temporary record is removed in WP-80. No runtime golden snapshot changed.

## WP-11 advisory-report verification

The native PIT extension now owns the two explicit module scopes. The custom
scope model, scope/report verifier tasks and their subject-only tests are gone.
All four mutation/coverage/test-strength/survivor thresholds are unconfigured,
as demonstrated by the convention-plugin functional test. The workflow has only
weekly/manual triggers and uploads contract and engine HTML/XML reports.

An independent temporary project deliberately asserted only that a conditional
return value was positive. PIT generated three mutants, killed one, and left two
survivors (33% mutation score/test strength). The run exited successfully in 7 s,
demonstrating advisory behaviour rather than merely inspecting configuration.
The probe's XML/log are retained outside the repository; its source/build tree
was removed. The actual `mutationReport` then passed in 49.26 s: contract killed
326/326 mutants and engine killed 82/82.

JaCoCo reports replace the percentage gates, including the separately discovered
72% gate in the nested Jazzer build. Reports now include that build's full main
source set. Test execution and committed-seed replay remain required.

The first configuration compile exposed two residual closing-call delimiters
after removing the old verifier blocks; they were corrected before testing.
The native task graph, formatting, convention-plugin test, shell contracts and
documentation contract passed. `./gradlew check --no-parallel` passed in 168.43 s;
`./gradlew coverage --no-parallel` passed in 162.23 s and produced the per-module
and aggregate HTML/XML reports. All 2800 tests pass with zero errors/skips,
including all 141 golden tests; no golden observation changed. Remote workflow
execution was not invoked.

## WP-12 scope challenge and independent WP-13 verification

Seven ordinary Jazzer test files and related fuzz consumers still import engine
internals. The bill explicitly schedules their model conversion in WP-40 and
their build ownership in WP-60. After the user again requested QA and proceeding
by engineering judgment, WP-12's interim import criterion is scoped to the
executor/test migration; the Jazzer consumers remain an explicit worklist for
WP-40/WP-60. This does not waive the final no-external-internal-import outcome.

WP-13 depends only on WP-10/WP-11 and therefore ran independently before WP-12,
as the bill permits for independent work within a phase. The legal-inventory
and JAR-pruning tasks are Kotlin implementations with the same Gradle property
getters and action/error behaviour. Existing legal-inventory controls remain;
two pruning tests protect declared archives, unrelated files, nested directories
and the missing-directory no-op. Main Java sources and Java-only compile cleanup
are removed. Repository helper functions have a responsibility-based file name,
and the included build derives its Java target from the root toolchain pin.

A clean focused rebuild passed all four task tests in 7 s. `javap` identifies
both action implementations' source files as `.kt` and retains their declared
IOException contract. `compileJava` reports `NO-SOURCE`. The final
`./gradlew check --no-parallel` passed in 271.58 s: 2802 tests, zero
failures/errors/skips, including all 141 golden cases. Packaging exercised the
ported tasks and no golden snapshot changed. All remaining Java under the
included build's `src` tree is test code.

## WP-12 test ownership and retained rules

`executor` is replaced by the verification-only `architecture-tests` project.
Runtime and parity suites now live in engine, including their resources and test
dependencies. Engine runs 1495 ordinary tests and all 14 parity tests. Architecture
verification has nine production checks and 11 focused negative controls for
their logical properties; JPMS owns the module graph. Rule-construction,
inventory, classifier-coverage and source-text architecture meta-tests are removed.
The compiled product roots come from the Gradle model rather than a parallel
anchor inventory. The verification project has its own JPMS descriptor; product
module exports were not widened.

API-signature checks now explicitly include POI/OOXML types and generics. Engine
violation fixtures belong to engine test fixtures; the verification suite inspects
their bytecode without importing engine internals. Current non-Jazzer ordinary
tests have zero such imports. The seven Jazzer test files remain the explicit
later worklist described above, together with their fuzz consumers.

The first moved-engine compile failed on seven missing Jackson tree symbols in
the moved fixture helper. Restoring its existing test dependency fixed the
consumer contract. Engine/parity tests then passed in 42 s. Architecture checks
and controls passed; focused PMD passed in 6 s. Updated shell/module/documentation
contracts passed. The final `./gradlew check --no-parallel` passed in 206.91 s:
2765 tests, zero failures/errors/skips, including all 141 unchanged golden cases.
The old ADR is explicitly historical pending its replacement in WP-20/WP-71.
