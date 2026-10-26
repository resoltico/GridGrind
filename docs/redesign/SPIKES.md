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

## Phase 1 whole-repository failure and focused repair

The first phase-end `./check.sh` passed quality/coverage (212 s), nested Jazzer
support and seed replay (32 s), and fat-JAR packaging (16 s). Stage 5 failed after
332 s because the grant-writing Python helper imported `typing.TypeAlias`, absent
from the ambient macOS Python 3.9.6. Total exit 1 after 607.43 s (script rounds
to 608 s); Docker was not reached in that run. The complete log is retained in
the package's `evidence/phase-1/check.log`.

The helper used a newer runtime typing expression without declaring that
prerequisite. Its recursive type descriptions now use standard `typing.Union`,
with no version switch, polyfill, dependency or grant-generation behaviour
change. The failing discovery verifier's focused test passes. Actual packaged
discovery execution passes every example and all eight task starters, and the
remaining container/publication contract tests pass.

Focused Docker smoke also passes (261.28 s): pinned-base platform inspection,
image build, legal visibility, packaged discovery, all recipe executions,
bind-mount user guidance, custom workdir/odd paths, existing-workbook reopening,
signature-line authoring under container fonts and streaming-write readback.
This built and removed only a local smoke image; no container was published.

The shell stage currently invokes the full quality gate again via its helper
inventory. That verified duplication is left for the scheduled WP-61 orchestration
replacement rather than bypassing the current canonical gate.

The repaired canonical `./check.sh` passed all six stages in 728.38 s (script
summary 728 s): hygiene, quality/coverage, nested Jazzer support and seed replay,
fat-JAR packaging, all shell contracts and actual Docker smoke. Docker's stage
passed in 121 s. The complete delivered-state log is retained as
`evidence/phase-1/check-repaired.log`. The earlier 607.43-second failure remains
recorded above and in its original log. No golden snapshot changed.

## WP-20 module consolidation checkpoint

The first reviewable WP-20 commit consolidates contract, Excel foundation and Java authoring
into `core`, preserving Java type identities and wire behaviour. This checkpoint deliberately
retains the codec dependency: WP-20 is not complete until the subsequent JSON extraction and
package normalization establish its final dependency rules. Docker, coverage, mutation, composite
fuzz consumption and current documentation now reference the consolidated project. The retired
three-project shape checker and obsolete architecture ADR were removed.

Production compilation passed after repairing a mechanically damaged CLI qualified `opens`
directive. A combined format/test invocation then triggered an internal Error Prone 2.50.0
`StringConcatToTextBlock` crash; isolated compilation passed with the check enabled, followed by
core tests and full `./gradlew check --no-parallel` (3m). The active suites contain 2,765 tests,
zero failures/errors/skips, including all 141 golden cases; no snapshots changed. Publication
contract regression passed. Evidence is external in `evidence/wp-20/`.

Regenerated inventories track the new core paths, avoiding false zeros from retired paths.
Main Java lines: 125,584 → 125,570 (redundant module descriptors); shell lines: 10,975 → 10,831.
Input/report/model/converter/union/catalog counts remain unchanged (81/96/51 snapshots/
13 definitions/4,576 converter lines/99 subtype files/2,849 descriptor lines).
`git diff main -- AGENTS.md` remains empty. Next: separate pure intake/problem data from JSON
parse state and remove databind from core before completing WP-20.

### WP-20 catalog template dependency removal

Catalog step templates now contain deeply immutable JDK maps/lists/scalars instead of Jackson
nodes. JSON conversion belongs to the output codec; CLI scaffold rendering requests a mutable
JSON copy there. Existing deterministic placeholder, executable round-trip and doctor tests are
preserved. A new regression changes nested source objects/lists after construction and attempts
mutation through published accessors, and rejects a mutable numeric object.

Focused catalog/doctor tests passed (26s). The first full check failed solely on eight PMD
`UseConcurrentHashMap` findings on local ordered builders. Narrow method suppressions document
thread confinement and immutable publication, matching existing catalog practice. Focused PMD
and template checks passed (14s); repaired full `./gradlew check --no-parallel` passed (2m44s),
2,766 active tests with zero failures/errors/skips, including unchanged 141 golden cases.
Evidence: external `evidence/wp-20/templates-*.log`. No golden snapshot was re-recorded.
Main Java lines 125,570 → 125,610; all model/converter/union counts remain unchanged.

### WP-20 pure intake/problem data ownership

Moved 47 problem, field-validation and partial-fragment types into `core.problem`, with their
pure tests and all consumers. Structural findings now expose their typed exception classification
in core; the JSON facade/decoder projection aliases were deleted. Engine classification no
longer calls the JSON codec. Parser-owned `RequestAnalysis` and binding normalization still need
the next extraction before the final module split. No Jackson import exists in `core.problem`.

The expanded move initially failed compilation on a duplicate export and package-private helper
references; a second focused compilation exposed missing Optional imports and a still-needed
private helper. Those were repaired without widening helper visibility. Focused behavioral tests
passed, but PMD required documentation on two newly public convenience constructors. Repaired
focused PMD passed (19s). Full `./gradlew check --no-parallel` passed (3m18s), 2,766 active tests,
zero failures/errors/skips, including unchanged 141 golden cases. Nested Jazzer formatting passed.
Logs and regenerated inventories are retained externally under `evidence/wp-20/`.

The private required-component constructor no longer repeats the blank-name check:
`GridGrindProtocolContractSupport.wireFieldName` already returns a nonblank annotation value or
its valid Java record-component name. Its null guard remains. This removes no authored-input
validation. Old JSON package files: 104 → 57; pure problem files: 47 plus package annotation.
Main Java lines: 125,610 → 125,704. Model/converter/union counts remain unchanged.

### WP-20 parser-independent intake handoff

Replaced parser-backed `RequestAnalysis` with the pure `core.problem.RequestIntake` record.
It retains independently bound siblings, ordered findings, optional complete plan, warnings and
an immutable canonical-path/token-offset index. JSON owns final binding, invariant failure
location and normalization. The old analysis object and binding-record normalization aliases
were removed. Engine main code now has zero imports of the JSON package. CLI redaction uses
`forRequestType(WorkbookPlan.class)`, exactly the metadata source previously selected inside
analysis; it never needed request values or parser state.

New regression controls independently assert the complete nested offset index, first-member
selection for duplicate keys, the reserved root location and ambiguous-name omission. A pure
handoff test checks offset ownership and rejects a complete plan alongside problems. Existing
binding, redaction, partial-sibling and source-location assertions remain. Focused PMD and intake
checks passed (6s); full `./gradlew check --no-parallel` passed (2m46s), 2,769 active tests,
zero failures/errors/skips, including unchanged 141 golden cases. Nested formatting passed.
Evidence: external `evidence/wp-20/intake-*.log` and `intake-inventories/`.

Old JSON package files: 57 → 56 (analysis plus binding data retired; binder added).
Main Java lines: 125,704 → 125,716. No golden re-recording, baseline or supplement change.
Next: physically separate JSON, move codecs out of core and normalize packages before WP-20 closes.

### WP-20 separate JSON module checkpoint

Created `json` with `codec` and `intake` packages. Core has no databind import/dependency;
its temporary Jackson annotations remain API-visible with `requires static` until WP-22.
Engine main has no JSON/Jackson imports or module requirement. Its last direct JsonProperty
lookup now delegates to the existing core metadata field-name function, preserving the same
nonblank-annotation/name fallback. JSON registers the selector serializer and step deserializer
in its mapper; core no longer carries their databind annotations. Java authoring emits plans,
with JSON output explicitly through the codec. JSON-dependent tests and fixtures moved with
serialization, preserving the total test count and package-private model assertions.

The initial split exposed 43 package-access compilation errors, followed by the engine's direct
annotation lookup and 17 test-access errors. Internal JSON package crossings were repaired;
private codec assertions moved into codec tests instead of widening those helpers. Redundant
JSON facade shortcuts and the test-only empty-redactor shortcut were deleted. The selector-hint
class had no production caller (only two assertions) and was removed early from the WP-23 list,
including those assertions. Its removal changes no runtime path. One focused test failed because
it still instantiated a bare mapper after registration moved out of core; it now uses the product
mapper and continues to assert the original missing-payload error.

All production/test compilation passed. Focused core/JSON PMD and tests passed except that
recorded mapper failure; its repaired focused rerun passed (4s). Full `./gradlew check --no-parallel`
passed (2m58s), 2,769 active tests, zero failures/errors/skips: core 328, JSON 310, engine 1,495+14
parity, CLI 596, architecture 11+9, build logic 6. All 141 golden cases passed unchanged.
Nested `jazzer check` passed (17s), including every regression target; publication-contract
regression passed. Core native PIT includes the moved JSON test source set and passed (11s):
298/298 mutants killed in its selected scope, no uncovered mutants. The old UTF-8 mutation target
is now JSON-owned and removed from core's scope. This is advisory reporting, not a score gate.

JAR inspection confirmed no retired `contract.json` classes in core, JSON or the packaged CLI,
and no serializer/deserializer implementation left in core. Core compilation emits eight
expected javac warnings for qualified opens to an optional databind module absent from its
compile path; the module's dependency was not restored to silence them. Docker, cleanup,
composite fuzz dependencies, docs and publication checks follow the five-module build. The
retired source assertions requiring codecs inside core were removed. Inventories track the new
module separately: old JSON package 56 → 0; JSON source files 63; main Java lines 125,716 → 125,668;
shell lines 10,831 → 10,827. Other model/converter/union counts remain unchanged.

Evidence: external `evidence/wp-20/json-*.log` and `json-inventories/`. No golden re-recording or
AGENTS baseline/supplement change. WP-20 remains open for core package normalization.

### WP-20 domain packages and completion

Core now uses `dev.erst.gridgrind.core` packages: vocabulary, metadata, feature models, plan,
source, problems/results, semantics, validation and catalog. Java authoring is in `core.authoring`.
Every caller, test, example, build scope and current source link follows the moved types. Engine
coverage now also collects JSON integration tests. The two distinct `NamedRangeTarget` contracts
(authoring target vs formula/range model) remain separate; a simple-name collision in the move
script initially left the model unmoved and was repaired with qualified mappings. Initial
production compilation had 81 access/missing-type errors; a descriptor-line repair then exposed
a dangling qualified-opens target, which was fixed by regenerating the descriptor from actual
packages. Shared existing validators retain their behavior across the new package boundaries.
Test-only private factories were exercised through equivalent public constructors or kept in the
owning security package; product APIs were not widened for those tests. Full production/test
compilation and focused core/JSON/architecture checks passed (17s).

The first full check failed (2m48s): making the optional open-security predicate public accidentally
exposed a JavaBean `empty` property in serialized requests. Two engine fixture tests, two CLI
fixture/materialization tests and both package-security golden transports detected it. Failure XML
is retained externally in `packages-failures/`. The plan now inspects the existing optional secret
reference directly; the unused convenience predicate and its redundant assertion were removed.
Focused security/fixture checks passed (53s), and repaired full `./gradlew check --no-parallel`
passed (3m2s): 2,769 active tests, zero failures/errors/skips, all 141 golden cases unchanged.
No snapshot was re-recorded. This producer/consumer failure is also relevant to WP-22's record-only
wire-property derivation: computed JavaBean predicates must not become wire fields.

Nested Jazzer `check` passed (19s), including all committed regression targets. `mutationReport`
passed (16s): core's selected 298/298 mutants and engine's selected 82/82 mutants killed, no
uncovered mutants. Publication regression passed. An independent Java 26 named-module probe
round-tripped all 14 shipped requests, the catalog and a response through the actual core/JSON
JARs. JAR inspection found no retired contract, authoring or Excel-foundation namespace in core,
JSON or the packaged CLI. Temporary annotation dependency and qualified opens remain until
WP-22/WP-24 remove the Jackson annotations/creators; no databind import/dependency exists in core.

Inventories now follow semantic packages instead of producing false zeros from the retired DTO
directory. The original DTO-origin input count remains 81, report count 96. A separate all-core
input count is 83, including the two pre-existing text/binary source unions omitted by the original
DTO-only measure. Snapshot/definition counts remain 51/13, subtype files 99, handwritten catalog
lines 2,849, blank checks 477, Support classes 220, JSON source files 63. Converter lines
4,576 → 4,575 reflect import/format changes only. Main Java lines 125,668 → 125,883; test lines
152,223 → 152,970; shell lines remain 10,827. The module consolidation began from six product/
verification projects; the current build has core, JSON, engine, CLI and architecture tests.

WP-20 acceptance is demonstrated by actual JPMS compilation and named execution, annotation-only
core's explicitly allowed temporary dependency, unchanged golden output, complete JAR contents,
and regenerated `package-inventories/`. Evidence is external in `evidence/wp-20/packages-*.log`,
`jpms-probe.log` and inventories. Baseline diff is empty; no supplement exists yet. Next: WP-21
value objects, then WP-22/WP-24 metadata/defaults and WP-23 vocabulary normalization.

## WP-21 validated scalar foundation checkpoint

Added sealed `WireString` and the eight required values: SheetName, CellAddress, RangeRef,
DefinedName, TableName, RgbHex, StepId and PlanId. They preserve authored spelling/endpoint order;
RGB retains the existing uppercase canonical form. Validation reuses the existing vocabulary
rules; RGB rules and their tests moved from style into vocabulary to avoid an upward dependency.
JSON has one generic constructor-backed scalar deserializer and one serializer, with no Jackson
annotation on the value records. This checkpoint does not yet retarget model/caller fields or
close WP-21: stringly inventories and duplicated guards still need to be drained.

Deterministic generated/edge cases exercise 1,049 valid values and 1,335 invalid values across the
eight types. Each valid value round-trips and its bytes are compared with the independently
expected primitive string (rather than an expectation read back from the value object). Invalid
values fail both constructors and codec binding; number/boolean/object/array tokens are rejected.
Optional omission and collection elements retain their scalar shape. One dynamic-factory PMD
false positive is narrowly documented: assertions run in each generated callback.

The initial focused run flagged throwing an unwrapped reflected cause; the codec now delegates
constructor failures to Jackson's instantiation handler, retaining the original domain cause.
A subsequent run exposed a missing-Optional bug: rejecting from `getNullValue` also rejects
absence through Jackson's Optional deserializer. The low-level codec now keeps the standard null
provider; explicit-null request rejection remains in intake, as D6 requires. Test quoting edits
also failed formatting/compilation and were repaired; all failure logs remain in `evidence/wp-21`.
Focused property/PMD checks passed (5s).

An independent oversized-row regression then failed with NumberFormatException on the existing
address validator. It now checks digit length before parsing, returning the existing BOUNDS
violation and domain message. The red log/XML and repaired focused run (15s) are retained. Bounds
constant names now correctly describe one-based row/column numbers. No supported valid address
or wire spelling changed.

Final `./gradlew check --no-parallel` passed (3m23s): 2,787 active tests, zero failures/errors/skips,
including all 141 unchanged golden cases. Nested Jazzer check passed (21s), including committed
regression replay. Native core PIT passed (11s), 298/298 selected mutants killed. The actual named-
module probe passed all 14 requests, catalog, response and all eight scalar values; vocabulary is
exported but not opened to databind, so the public validating constructors work without private
reflection access. No snapshots were re-recorded and no runtime dependency was added.

Inventories: blank checks 477 → 479 (new PlanId/StepId constructors before caller migration),
main Java lines 125,883 → 126,109, test lines 152,970 → 153,173, JSON files 63 → 66. Model/mirror/
subtype/catalog counts are unchanged. Evidence: `value-inventories/` and `values-*.log` externally.
Baseline diff is empty; no supplement change. Next: teach scalar metadata/intake consumers about
WireString, migrate identifiers across model/engine/hosts, and delete the superseded guards.

### WP-21 scalar metadata and plan identity

One `WireString.isValueType` predicate now drives catalog scalar shapes, request-object graph
traversal, tolerant token-kind validation, secret-path traversal, source-origin traversal and the
scalar codec. Domain values stay strings rather than acquiring a synthetic `.value` object path.
Tests cover every permitted scalar shape, explicit-null Optional input, and exclusion of scalar
records from the request-object graph.

`PlanId` now flows through plans, partial fragments, Java authoring, result factories/results and
CLI example/starter construction. Duplicate blank checks and result-plan-id normalization were
removed; canonical value construction owns validity. Container null guards remain. No String
constructor alias was added. Tests compare literal expected ID text through `.value()`; an
existing materialized-example self-comparison was corrected to compare against the authored
request. New live intake tests check omission, string serialization, invalid values and exact
path/token locations.

A temporary compiler-informed argument tool outside the repo resolves actual qualified symbols
across 2,020 source files with the 40 resolved library JARs, then rewrites only String-to-domain
arguments/initializers. It reports no edits on the unchanged baseline. It also handles empty/poly
factory inference rather than treating standalone inferred Object collections as incompatible.
The first classpath collector hit Gradle's cross-project resolution lock; per-project collector
tasks fixed that. No product build task or dependency was added for the tool.

Initial Gradle compilation identified two obsolete self-assignments after removing normalization;
those now use ordinary non-null checks. Focused runs then found four expected String-vs-PlanId
assertion mismatches and an implementation-name wrapper in the new invalid-ID diagnostic. The
codec now emits an owned invariant for expected untyped constructor failures, retaining the
original cause and preserving already typed problem sources. No Jackson-message regex was added.
The new scalar properties assert that owned failure type; non-string token failures remain codec
shape errors. Repaired focused checks passed (28s).

Full `./gradlew check --no-parallel` passed (2m56s): 2,792 active tests, zero failures/errors/skips,
including all 141 unchanged golden cases. Nested Jazzer check passed (17s). `mutationReport`
passed (21s), selected core 298/298 and engine 82/82 mutants killed. Actual named-module request,
catalog, response and eight-value probes passed. No golden re-recording, baseline or supplement
change. Evidence: external `plan-id-*.log`, `plan-id-failures/`, `plan-id-inventories/`.
Blank checks 479 → 478; main lines 126,109 → 126,145; test lines 153,173 → 153,289. Other counts
remain unchanged. WP-21 remains open for step and workbook identifiers and obsolete-guard removal.

### WP-21 step identity throughout execution

StepId now flows through plan steps, authored Java steps, read commands/results, assertions,
journals/progress, warning/problem contexts, host acceptance and semantic preservation evidence.
The compiler-informed tool updated 908 String arguments/initializers in 107 files, followed by
83 typed-result-to-literal-text projections in 34 assertion files. Expected literal text remains
independent of constructor behavior. Primitive step checks were removed from containers, and
WorkbookStepValidation was deleted; its behavioral operation-target tests remain. String labels
remain at textual diagnostic boundaries that can represent either step IDs or OOXML part names.

Production compile repairs covered the failed-step guard and typed sets/lists. Test compile repairs
covered fixture lookups/text projections, three authoring fixture IDs and typed preservation IDs.
Focused core/JSON/engine/parity tests passed except the missing LIM-006A marker; it now belongs to
the canonical StepId validator, with LIMITATIONS and Java authoring guidance updated. Live request
tests assert scalar encoding plus invalid-ID message, authored path and token offset. Focused
checks passed (15s). One recipe crossed PMD's import threshold because of the added domain type;
a narrow documented exception keeps that cohesive cross-feature recipe intact (focused 6s).

The next full check exposed a separate CLI default mapper unable to read typed progress IDs.
CLI-owned binding now reuses the canonical configured mapper, and its redundant non-closing stream
wrapper was removed (the mapper already disables closing caller streams). An added grant acceptance
round trip then found that the old custom decoder was registered only for WorkbookStep, not the
concrete assertion/inspection types used by host grants. Step codecs are now registered from the
sealed permits list, and the generic decoder checks the expected concrete kind. A negative grant
control rejects an inspection in a terminal-assertion slot. Focused progress/grant/JSON/PMD checks
passed (16s). Failure XML is retained as `step-id-progress-mapper-failure.xml` and
`step-id-concrete-grant-failure.xml` externally.

Final `./gradlew check --no-parallel` passed (2m37s): 2,796 active tests, zero failures/errors/skips,
all 141 golden cases unchanged. Nested Jazzer support tests initially found two internal read-probe
IDs formed from sheet labels with spaces; valid fixed probe IDs now leave labels and compared
facts untouched. Focused seven-case round-trip verification passed (8s), then full nested `check`
passed (10s), including committed replay. `mutationReport` passed (24s), selected core 298/298 and
engine 82/82 mutants killed. Named-module request/catalog/response/eight-scalar probes passed.
No golden re-recording or baseline/supplement change. Evidence: external `step-id-*.log`, failure
XML, `step-id-jazzer-failures/` and `step-id-inventories/`.

Blank-check sites 478 → 474; main lines 126,145 → 126,314; test lines 153,289 → 153,687. Converter
lines 4,575 → 4,576 are import/format changes only. Other model/mirror/subtype counts are unchanged.
WP-21 remains open for workbook identifiers and the remaining obsolete guards.

Before that name migration, an independent malformed-workbook probe changed one existing sheet
label to `invalid[]` without changing workbook facts. The current summary operation succeeded and
reported that label. Authored selectors already validate Excel names; observed report labels have
a distinct, wider contract. Preserve that diagnostic capability with a distinct observed-label
value instead of applying authoring-name validity to observed source facts. Evidence is external
in `observed-names/` (owned workbook copy, request/grant/result and run log).

### WP-21 observed worksheet labels

The malformed-workbook probe established a real distinction: authored sheet names are validated
against Excel naming rules, while workbook summaries already report wider observed labels such
as `invalid[]`. Added SheetLabel for that observed contract; SheetName retains its authoring
rules. WorkbookSummary and SheetSummaryReport now carry typed labels, with scalar JSON output
unchanged. Converters construct labels at the observation boundary. The collection helper now
has the accurate generic name `copyDistinctValues`: it copies nonnull values and rejects duplicates
(it never removed duplicates). Count, active/selected membership, nonempty-selection and null-
element invariants remain; repeated primitive blank checks moved into the label constructor.

The compiler tool adapted immutable literal lists and typed text assertions. Four parity compile
checks and one Jazzer invariant needed explicit `.value()` projections against the independent
POI/string oracle. One constructor assertion still expected the retired container-level blank
message; it now checks the canonical label message. Repaired focused checks passed (12s).
A persisted-workbook regression bypasses POI's authoring setter, writes `invalid[]`, then executes
product summary inspection and response round-trip to prove that observed labels remain readable.
The generated scalar suite also covers SheetLabel's wider contract and invalid blank values.

Full `./gradlew check --no-parallel` passed (2m43s): 2,799 active tests, zero failures/errors/skips,
all 141 golden cases unchanged. Nested Jazzer check passed (16s) after the textual invariant
projection. `mutationReport` passed (15s), selected core 298/298 and engine 82/82 mutants killed.
Named-module probes passed all 14 requests, catalog, response and nine scalar values, including
an invalid authoring-name label. No golden re-recording or baseline/supplement change.
Evidence: external `sheet-label-*.log`, `sheet-label-failures/`, `sheet-label-inventories/` and the
original `observed-names/` probe. Blank checks 474 → 473; main lines 126,314 → 126,344; test lines
153,687 → 153,826; converters 4,576 → 4,583 (boundary scalar construction, mirrors still pending
WP-40). Other model/subtype counts remain unchanged. Next: authored sheet identifiers.

### WP-21 authored sheet identifiers

Authored selectors, actions, fluent targets, table/pivot/calculation payloads and named-range
scopes now carry `SheetName`. Backend adapters and diagnostic projections extract text at their
boundaries. Case-insensitive selected-sheet duplicate detection remains. The obsolete plan-layer
sheet-name validators and their forwarding assertion were deleted; LIM-003 now marks the value
constructor and the limitations reference points to it. Shipped Java examples, recipes and fuzz
producers use the typed API.

The first full check detected two regressions: the compile-verified Java example still supplied
strings, and an invalid selector name lost its precise `target.name` path. The scalar codec now
lets Jackson retain field/index references around the owned validation cause, and step failure
projection retains that stronger path. The generated-invalid properties independently assert the
owned cause, and an invalid second cell-address regression asserts both the field and array index.
Focused example/property/CLI checks passed after repair. Nested Jazzer compilation also detected
ten comparisons of typed authored names with textual observed facts; these now compare their
text, preserving the independent response checks. Nested check plus committed replay passed (18s).
All earlier failures remain in external logs/XML.

Full `./gradlew check` passed (2m52s), followed by the final check after the limit-marker/comment
and documentation correction (3m26s): core 328, JSON 338, engine 1,510, CLI 598, architecture 35;
2,809 tests total, zero failures/skips. All 141 golden cases passed without snapshot changes.
Named-module probes passed 14 requests, catalog, response and nine scalar types. Inventories:
blank checks 473 unchanged; main Java 126,344 -> 126,476; test Java 153,826 -> 154,276; converter
lines 4,583 -> 4,601. Inputs 81 (83 including source unions), reports 96, snapshots 51, definitions
13, subtype files 99, hand-written catalog lines 2,849, JSON files 66 and Support files 220 remain
unchanged. Evidence: external `sheet-name-*.log`, `sheet-name-failures/`, `sheet-name-inventories/`
and compiler-informed migration logs. `git diff main -- AGENTS.md` remains empty; no supplement.
WP-21 continues with observed sheet labels and the remaining value-typed fields.

### WP-21 observed sheet fields and diagnostic locations

Fourteen report/result model files now use typed sheet fields: `SheetLabel` for factual labels,
`SheetName` for the already-strict pivot report/source contract. Problem/warning locations and
their sheet accessors carry `SheetLabel` as well. No core `String sheetName` or
`List<String> sheetNames` declaration remains. Removed repeated label blank checks and the unused
inspection-result text validator. Textual fuzz/parity or backend consumers project `.value()`.

The malformed-sheet regression now persists a table before renaming its sheet to `invalid[]`;
reopened workbook-summary and table reports preserve that label and response round-trip. During
review, a codemod had copied a mutable test list before the constructor under test; the test now
uses a caller-owned mutable `List<SheetLabel>` directly, retaining its defensive-copy assurance.
The defaulting test's bare Jackson mapper could not bind scalar values; switching it to the
canonical codec exposed stale journal fixture fields that its old permissive mapper ignored.
The current `ExecutionJournal` has no planId/persistence/warnings fields and `Phase.NotStarted`
has zero components. Removed those unknown fields, retaining valid outcome timing. Focused tests
pass. Retained both failure XMLs in external `observed-sheet-failures/`.

The first full check caught two parity comparisons of text with typed labels and standard PMD's
30-import threshold in two explicit DTO adapters (30 and 31 imports). Corrected comparisons;
scoped `PMD.ExcessiveImports` suppressions document the intrinsic backend/report correspondence,
without changing the ruleset or fragmenting the adapters. These converters remain WP-40 deletion
targets. Focused compilation, PMD and all 14 parity scenarios passed (27s).

Final `./gradlew check` passed (3m13s): core 328, JSON 338, engine 1,510, CLI 598, architecture 35;
2,809 tests, zero failures/skips. All 141 golden cases unchanged. Nested Jazzer check and committed
replay passed (25s); named modules round-tripped 14 requests, catalog, response and nine scalar
types. Inventories: blank checks 473 -> 462; main Java 126,476 -> 126,506; test Java 154,276 ->
154,463; converters 4,601 -> 4,614 lines. Other model/subtype counts unchanged. External evidence:
`observed-sheet-*.log`, `observed-sheet-failures/`, `observed-sheet-inventories/`, compiler migration
logs and the classified file list. No golden re-recording, baseline or supplement changes.
Next: cell-address fields; retain typed address remediation and range-selector endpoint ordering.

### WP-21 cell addresses

Fourteen model/authoring files now carry `CellAddress`: cell selectors and lists, window anchors,
cell/readback reports, formula targets and pivot anchors. Constructor validation uses the shared
address problem mapper, retaining typed rules and remediation with no invented owner path.
Case-insensitive address/name list copying shares one typed helper; the read count limit remains.
Removed the obsolete list address-format branch and its unreachable helper-only assertion.
The resolver's private numeric row/column pair is now `CellCoordinates`, avoiding confusion with
the wire value. Invalid diagnostic tokens remain text; backend/rendering/parity/fuzz consumers
extract scalar values at their boundaries.

Focused tests caught old list-index/direct-topLeftAddress message expectations: the value now
names `address`, while the public intake owns field/index location. A public request regression
asserts `steps[0].target.addresses[1]` and the syntax rule; nested codec regression asserts the
bounds rule and array path. The doctor out-of-bounds golden message/remediation/location remain
byte-identical. A full check exposed synthetic `A2b`/`A2c` addresses in a facet-conversion fixture;
replaced them with distinct valid `A8`/`A9`, preserving all facet assertions. Mixed text/value
comparisons and a parity map projection were corrected. Standard PMD's import threshold required
one scoped, documented `ExcessiveImports` suppression on the semantic selector resolver (30
imports); its selector/backend input contract is cohesive. No ruleset changes or new fragments.
Earlier failures retained externally, including XML in `cell-address-failures/`.

Final `./gradlew check` passed (3m02s): core 328, JSON 339, engine 1,510, CLI 598, architecture 35;
2,810 tests with zero failures/skips and all 141 unchanged golden cases. Doctor-only corpus passed
independently (12s). Nested Jazzer check and replay passed (23s), focused parity all 14 cases, and
named modules round-tripped 14 requests, catalog, response and nine scalars. Inventories: blank
checks 462 -> 451; main Java 126,506 -> 126,567; test Java 154,463 -> 154,832; model/subtype counts
unchanged. Evidence: external `cell-address-*.log`, inventories, failures, classified file list and
compiler-informed migration logs. Baseline diff remains empty; no supplement or golden re-record.

Range follow-up QA: a direct reopened-workbook experiment stores `sqref="#REF!"` on a future-type
validation. Current product exits 0 and reports `UNSUPPORTED` with ranges `["#REF!"]` truthfully.
Evidence: external `range-observations/` Java producer, workbook, request/grant and result/logs.
An authored local `RangeRef` cannot represent that observed fact. Preserve it with a nonblank
factual range-label value, alongside strict authored ranges. Keep selector-specific endpoint
ordering (the backend rectangle parser normalizes reversed endpoints). Print-area readback was
also inspected: it normalizes stored formulas to local rectangular references; its authoring
path parses one rectangle, so no speculative print-area expression abstraction is needed.

### WP-21 authored and factual range fields

Authored model/authoring range fields now carry `RangeRef`; factual report fields carry
`RangeLabel`, preserving raw malformed references established by the reopened-workbook probe.
Both serialize as strings. Added generated valid/invalid range-label properties (including
`#REF!`, qualified and future labels), and a persisted unsupported-validation regression proving
its stored `#REF!` survives readback and complete response round-trip. The named-module response
probe now includes that fact; ten scalar types are exercised.

`RangeRef` owns syntax/bounds validation with structured range/address rules. Selector ordering
remains a separate invariant, checked before duplicate detection in each range-list entry.
Removed obsolete selector address/range-format helpers and their deleted-branch assertions;
public value properties and typed syntax/remediation checks exercise the replacement. Typed
mutable input lists preserve defensive-copy test assurance. Updated Java examples, recipes,
fuzz producers and independent parity text projections. No aliases or snapshot re-recording.

A red test confirmed that a valid anchor plus positive, size-bounded dimensions could extend past
the last Excel row/column. Window construction now rejects this inconsistent combination and
retains the valid `XFD1048576` 1x1 case; computed windows return `RangeRef`. The red test initially
needed its missing assertion import fixed; `ranges-window-red-confirmed.log` and the retained XML
show the actual intended failure before the production correction. This is a coordinate-domain
invariant, not a new size policy. Field/count validation order and existing read-cell limits stay.

Focused core/codec tests, all 14 parity cases and the persisted-range regression passed (45s).
The first compile/check pass exposed five text/value parity comparisons and one mutable-list
input, corrected without changing expected facts. Standard PMD's threshold also required a
scoped `ExcessiveImports` suppression on the complete dashboard recipe (30 imports), documented
in place; no ruleset changes or recipe fragmentation. Evidence retained in `ranges-failures/`.
Nested fuzz support and committed replay passed (20s) after projecting factual range text and
constructing typed conditional inputs; model constraints were not caught or bypassed.

Final `./gradlew check` passed (3m20s): core 329, JSON 341, engine 1,511, CLI 598, architecture 35;
2,814 tests, zero failures/skips; all 141 golden cases byte-identical. Named modules passed 14
requests, catalog, response with malformed-range evidence and ten values. Inventories: blank
checks 451 -> 440; main Java 126,567 -> 126,643; test Java 154,832 -> 155,129; model/subtype counts
unchanged. External evidence: `ranges-*.log`, inventories, failures, red test XML, range field list,
compiler migration logs and `range-observations/`. Baseline diff empty; no supplement.
Next: RGB fields and remaining names. Optional fields follow the adopted `Optional<T>` rule
(target architecture section 4); a nullable data-bar color needs that explicit representation.

### WP-21 RGB ownership and optional factual colors

Style RGB fields and color-scale lists now carry `RgbHex`; data-bar RGB is `Optional<RgbHex>`.
The adopted vocabulary explicitly requires optional fields to use Java Optional and be absent
on the wire (target architecture section 4). Constructor validation/uppercase normalization is
owned once by `RgbHex`. Removed `ProtocolRgbColorSupport`, its verification-only constructor
factory and unused null-normalizer assertion. Meaningful helper tests now exercise the actual
value constructor in `RgbHexTest`; codec regression covers absent/present report colors and
round-trip. Independent parity/fuzz checks project scalar text. A factual absent color was
already allowed by the old model; fuzz validation now checks the optional wrapper and present
values instead of demanding a String value that the model allowed to be absent.

PIT selection now targets the replacement class/test and the existing value property test.
A fresh configuration uncovered an unsafe cross-project `json:testRuntimeClasspath` resolution
while serializing PIT's foreign test SourceSet into Gradle's configuration cache. The first
owned-variant attempt supplied a test JAR: it ran but missed JSON test discovery (core 162/294,
132 uncovered: 131 catalog mutants plus one value toString). A class-directory variant restored
discovery but omitted JSON production classes, yielding 28 baseline test failures. Adding the
production classes to the same owned dependency runtime restored the intended coverage. The
final JSON variant exports main/test class directories with their runtime dependencies, and
core's native PIT task resolves its own configuration. No foreign runtime configuration is
resolved directly. No test skipping, threshold or report-exclusion changes.

Final advisory mutation execution passed (15s): core 294/294 killed, engine 82/82 killed, zero
uncovered mutants. The four-mutant reduction from the previous core set reflects retirement of
the optional RGB helper. RgbHex constructor and toString mutations are covered by the real value
unit/property tests. Configuration-cache reuse passed (415ms). Earlier failure/discovery logs
retained (`rgb-mutation*.log`), final mutation XML copied to external `rgb-mutation-results/`.
The abandoned test-JAR task was removed and jar pruning left only the normal JSON artifact.

Focused core/codec/parity/PMD checks passed; a stale direct String expectation for the now-optional
report color was corrected without changing the expected color. Discovery goldens passed
independently (14s); nested fuzz support and replay passed (18s). Full `./gradlew check` passed
(3m08s), and the final documented/staged state passed (3m04s): core 329, JSON 342, engine 1,511,
CLI 598, architecture 35; 2,815 tests, no failures/skips, all 141 golden cases unchanged. Named
modules passed 14 requests, catalog, a response containing observed `#REF!`, and ten values.
The retired RGB helper is absent from the core JAR. Staged inventories correctly exclude its
retired file: Support classes 220 -> 219; blank checks 440 unchanged; main Java 126,643 -> 126,622;
test Java 155,129 -> 155,238. Other model/subtype counts unchanged. Evidence external
`rgb-*.log`, failures, inventories and migration logs. Baseline diff empty; no supplement.
Next: defined/table identifier names, preserving factual-name contracts separately where needed.

### WP-21 defined names and observed labels

Named-range selectors, mutations, authoring targets and pivot named-range sources now carry
`DefinedName`; the constructor retains structured naming failures through the common mapper.
Case-insensitive selector identities and name-list duplicate rules remain. Deleted obsolete
mutation name validators/forwarders and the validation helper's verification-only construction
factory/assertion; its meaningful grammar tests remain. Backend/diagnostic/parity/fuzz consumers
project text at their boundaries. Core/codec tests initially exposed two stale text/value
assertions, corrected with independent literal expectations; compiler checks also caught four
mixed parity/runtime-test comparisons. Failure evidence retained in logs/XML.

A direct external reopened-workbook probe establishes two real factual contracts: a defined name
`1invalid` is successfully reported, and GET_TABLES reports a stored empty table name. The first
producer attempt accessed CTDefinedNames before POI serialized its logical name list; the
repaired producer modifies workbook.xml after writing and preserves the exact original formula.
Evidence: `name-observations/` source, creation logs, workbook, request/grant and successful result.
`DefinedNameLabel` preserves the former capability with the old nonblank-only factual contract;
its generated valid/invalid/codec properties include invalid authoring syntax and reserved names.
A persisted product regression changes the stored name through ZIP filesystem XML, reopens it,
asserts `1invalid` and `Data!$A$1`, and verifies response round-trip. Table labels follow next.

Final `./gradlew check` passed (3m34s): core 329, JSON 344, engine 1,512, CLI 598, architecture 35;
2,818 tests, zero failures/skips; all 141 golden cases unchanged. Focused core/codec/persisted-name
and all 14 parity checks passed; nested Jazzer support/replay passed (19s). Named modules passed
14 requests, catalog, response and eleven scalar values. Inventories: blank checks 440 -> 439;
main Java 126,622 -> 126,670; test Java 155,238 -> 155,391; Support classes 219 unchanged, other
model/subtype counts unchanged. External evidence `defined-name-*.log`, inventories, failure XML,
migration logs and the name observation probe. No baseline/supplement or golden snapshot changes.
Next: table-name values and factual empty table labels, then the remaining WP-21 inventory.

### WP-21 table identifiers and factual empty labels

Table selectors, table specs, pivot table sources and fluent table identities now carry
`TableName`. Shared Excel-name validation produces the structured problem for both authored
name value constructors; removed duplicate primitive rendering and obsolete selector validators.
The initial broad Tables edit also changed style names; compilation caught that distinct
contract, and style fields/factory calls were restored before delivery. Table styles are not
table identifiers. Independent parity and diagnostic consumers project scalar text as needed.

The successful reopened-name probe established that GET_TABLES reports an empty stored table
name. `TableLabel` preserves exactly the former nonnull-only factual contract, including empty,
whitespace, malformed and Unicode strings; constructor null and non-string wire tokens are the
invalid domain. Generated properties reflect that contract rather than inventing syntax rules.
A persisted product regression writes an empty CTTable name, reopens it, asserts the empty
label, and verifies response round-trip (so the required empty String is not silently omitted).
Updated Java examples and typed list tests; no compatibility overloads or aliases.

Focused checks caught stale scalar/list parity/test expectations; all were corrected while
retaining independent literal facts and mutable-input copy assurance. Fuzz conditional input
construction now uses TableName; nested support/replay passed (21s). Native mutation report
passed (18s): core 291/291 killed, engine 82/82 killed, none uncovered. The smaller selected
core set reflects deleted primitive rendering and verification-only factory branches. Named
modules passed 14 requests, catalog, response and twelve scalars.

The first full check found one stale CLI table-name assertion; focused repair passed (6s).
Final `./gradlew check` passed (2m58s): all tests green, all 141 golden cases unchanged. Inventories:
blank checks 439 -> 438; main Java 126,670 -> 126,665; test Java 155,391 -> 155,574; Support classes
219 unchanged, other model/subtype counts unchanged. External evidence: `table-name-*.log`,
inventories, failure XML, migration logs and `name-observations/`. Baseline diff empty; no
supplement. The named-range surface report and other remaining domain names still need review
against the WP-21 inventory; this is not completion of WP-21 or the foundations redesign.

### WP-22 metadata identity/catalog checkpoint (WP-24 still pending)

Replaced colocated ProtocolTypeMetadata and every JsonSubTypes list with core Union/Variant
identity. CatalogType owns standalone-record descriptions, ProjectedBy owns facet-gated fields,
and FieldRule owns the six scalar constraint kinds. Runtime ModelGraph traverses request,
result, catalog and progress roots; no classpath scan or handwritten type inventory remains.
CatalogTypes derives descriptors and group ids, CatalogGroups resolves their field references.
Deleted manual descriptor files, the duplicated coverage validator, and field-shape registry.
Presentation order is explicit model metadata; source permits order preserves subtype order.
Shared pattern/bound constants remain referenced, rather than copied into annotations. Selector
identity also derives from the sealed graph, and SelectorTypes replaces SelectorJsonSupport.

Operation Variant metadata selects one of the shared conservative effect profiles and declares
preconditions. An independent comparison against the captured pre-change map shows effects,
footprints and precondition classes unchanged for all 126 operation records. Catalog constraints
previously included a name entry on the fieldless NamedRangeTarget interface: it could never
publish a real component rule, so the derived model drops that invalid registry-only entry.
Actual mutation/table record constraints remain. Retired registry/Jackson-fallback tests were
removed; replacement metadata identity, real consumers, complete golden corpus and architecture
negative controls exercise the current paths. Declaration-order diagnostics were restored after
a focused engine assertion detected an accidental allowed-id sort change.

Two brief premises required correction under the user's AFK QA delegation:
1. Three named-range records had different summaries in two catalog groups. One record now owns
   one canonical reference summary. The public structured catalog changes those three texts.
2. The nested pivot Source unions collided under the prescribed simple-name group algorithm.
   Their planned vocabulary roles were moved forward: PivotSourceSpec and PivotSourceReport.
Closed sealed subinterfaces remain, rather than flattening them into independent facets:
a Java 26 compile experiment demonstrated loss of exhaustive-switch coverage in that approach.
The architecture rule now targets Union and enforces record leaves plus closed intermediates;
its enum/open-family negative control remains. This explicitly deviates from direct-permits-only
records while preserving the stronger existing subset contracts and compiler checks. Evidence:
`facet-compile.log` and prototype. Do not claim direct-record-only permits in final reporting.

Reviewed golden changes are exactly ten files. Discovery 10 changes 33 group ids without member,
count or presentation-order loss. Discovery 11 changes those identifiers and consequent
identifier-based tie/support order. Eight recipe detail snapshots change capability group refs
only: TABULAR_REPORT, DASHBOARD, DATA_ENTRY_WORKFLOW, PIVOT_REPORT, AUDIT_EXISTING_WORKBOOK,
CUSTOM_XML_WORKFLOW, DRAWING_AND_SIGNATURE_WORKFLOW, WORKBOOK_MAINTENANCE. Both launchers' failure
observations equal the original snapshots after group-id mapping; behavioral facts remain fixed.
No other snapshot changed. Review evidence `catalog-*-review.json`, `recipe-group-id-review.json`,
full failure XML and `group-id-changes.json` is external. Registry/API documentation, fixture
comments and stale group-reference assertions were corrected from full failure output; retained
logs explain all intermediate failures.

Final ./gradlew check passed (4m20s), including all 141 golden cases, new metadata consumers and
nonempty Union architecture checks. Named modules passed 14 requests, catalog, malformed-range
response and twelve scalars. Inventories: registrations 5 -> 0, handwritten descriptor lines
2,849 -> 0, Support 219 -> 217, blank checks 438 unchanged, main Java 126,665 -> 125,348, test Java
155,574 -> 154,930; engine mirrors/converters remain scheduled work. Remaining Jackson creator,
property and inclusion annotations in core must still be removed with WP-24; the placeholder
WireDefault annotation is not yet used. This checkpoint does not complete WP-22 or WP-21.
No mutation/active fuzz campaigns were run after the user's updated instruction; defer them until
all planned implementation is finished, and cap any justified interim run at 120 seconds.

Defaults mechanism prototype passed outside the repository: absent Boolean default, explicit
false, Optional absence, explicit-null rejection, and default omission through native Jackson 3
value-instantiation/serialization extension points. See `default-binding-and-omission-prototype.log`.
Next: apply metadata defaults, delete secondary wire creators/null normalizers, remove the
annotation dependency from core, then complete vocabulary normalization and schema intake.
Baseline diff remains empty; no supplement.

### WP-24 omission-default mechanism and execution policy checkpoint

WireDefault now declares literal, public factory, and sibling-derived component defaults. A pure
JDK resolver supplies Java values; native Jackson 3 creator/property extensions distinguish absent
fields from explicit null and omit only declared matching defaults. NON_ABSENT is global. Execution
policy, calculation, journal and formula-environment secondary wire factories were deleted, with
seven obsolete custom inclusion filters and three tests dedicated to those deleted filters. Five
replacement tests independently assert default/override/null behavior, nested policies, sibling
resolution and output omission. Boolean sites carry WireDefault while old consumers are migrated in
the next batch; this checkpoint does not complete WP-24 or remove the core annotation dependency.

Core and JSON regression suites passed. Initial compile errors (NullAway) and standard PMD findings
were repaired; focused PMD then passed. One overlapping iteration failed while another build pruned
JAR outputs; stable checks were rerun serially. A combined spotlessApply/check invocation exposed
an existing task-order cycle, so formatting and verification were run as separate invocations.
The stopped Gradle daemon had retained hundreds of idle compiler workers; terminated that stopped
daemon and its task-owned workers before the final build. No test result from the overlapping runs
was used as delivered-state evidence.

Final ./gradlew check passed in 3m20s: core 325, JSON 335, engine 1513, CLI 598, architecture 35;
no failures, errors or skips. All 141 golden cases passed without snapshot changes. Inventory:
main Java 125348 -> 125580, test Java 154930 -> 154931, JSON main files 67 -> 69; registrations 0,
handwritten descriptor lines 0, Support 217 and blank checks 438 unchanged. Evidence is external in
wp-22/defaults-*.log and defaults-inventories. No mutation or active fuzzing ran. Baseline diff is
empty and no supplement exists yet. Next: complete all defaults/creator and Jackson-core removal.

### WP-22 / WP-24 canonical records and published defaults

Core is now Jackson-free: removed all creator/property/include/ignore annotations, its annotation
API dependency, static module requirement and all 21 Jackson-specific opens declarations. Closed
union identity remains metadata-owned. Secondary wire creators and ProtocolBooleanDefault were
removed throughout core and the last CLI creator was replaced by canonical status state.
Record components now own requiredness and wire names; aliases/creator-shape reflection were
retired. Required primitive omission/null rejection remains enforced by the JSON codec. Java
WorkbookPlan construction now requires explicit non-null policies; wire omission uses metadata.
Public table default factories replace inaccessible package-private factories. WireDefault supports
only the concrete needs found here: literals, no-argument public factories and sibling-derived
values; unused factory-argument scaffolding was removed.

The catalog replaces defaultBoolean with defaultValue and defaultFrom, publishing enum, string,
boolean, empty-list, source-object, record and sibling defaults from the same annotations. Published
value trees defensively copy mutable containers. Tests independently assert actual default values,
derived-field metadata, defensive copying, omitted/explicit/null inputs and nested policies.
Former creator/alias-only fixtures were removed; real presence and primitive-boundary checks were
retained against canonical metadata. Source-origin tests now protect canonical component names.
An architecture rule rejects POI/Jackson core dependencies, with a negative fixture detecting both;
Jackson is test-scoped in architecture-tests solely for that negative control.

Named-module verification passed without qualified opens: 14 request round trips, catalog, response
and twelve domain strings. Intermediate runs exposed a required-marker Boolean/null unboxing bug,
a package-private default factory, untyped fixture serialization dropping union tags, stale alias
fixtures, and unused JPMS requirements. Each was corrected and checked in the affected suite before
full verification. The missing-status rejection test now removes the actual JSON property rather
than relying on comma position. Named logs retain the outcomes; retired-origin-alias-test-failure.xml
retains the integration failure for the removed behavior.

Golden review: exactly one snapshot changes, discovery/10-print-protocol-catalog.json. It adds two
fieldMetadataKeys descriptions for defaultValue and defaultFrom, with no other content change.
Comparison with the prior committed snapshot and packaged-JAR output proves equality after removing
those two entries (defaults-catalog-index-review.json). Both launchers' discovery tests pass. The
initial snapshot-update command continued after an incorrect JAR filename assumption; the complete
diff was subsequently checked against the prior commit and actual packaged output. No unexpected
snapshot changes were retained. The final full run passes all 141 golden cases, including the
behavior-changing negative control, without further recording.

Final ./gradlew check passed in 3m47s: core 328, JSON 333, engine 1511 (including parity), CLI 598,
architecture 37; no failures, errors or skips. Inventory from the prior checkpoint: main Java
125580 -> 124461; test Java 154931 -> 154570; creators/default sites -> 0; registrations and manual
descriptors remain 0; Support 217, blanks 438, engine snapshots 51, definitions 13, converter lines
4695 remain. Acceptance evidence is external in defaults-delivered-inventories and defaults-*.log.
Baseline diff is empty; supplement absent. No mutation or active fuzzing ran.

WP-22's core/Jackson and derived catalog requirements now hold. WP-24's creator/default inventories
are empty; schema default publication is coupled to the scheduled WP-30 schema implementation and
remains outstanding until that consumer is verified. WP-21's full identifier/blank-check target is
still outstanding; do not count the checkpoint as its completion. Next: WP-23 vocabulary, then
schema intake and the remaining engine work; campaigns stay deferred until planned implementation
is complete, with any justified interim campaign hard-capped at 120 seconds.

### WP-23 authored-model naming checkpoint

Renamed 83 authored top-level Input types to 75 Spec and eight Patch types, plus the nested
CustomConditionSpec and transport Stdin names. Patch names are reserved for partial, optional-field
updates with nonempty-constructor invariants; complete policy/definition choices use Spec.
CellInput is now CellValueSpec, and ConditionalFormattingDefinitionInput is ConditionalFormattingSpec.
There are no aliases. Java callers, reflective tokens, catalog ids/group references, current docs,
scripts and fuzz source callers were migrated together. Historical CHANGELOG and redesign records
were preserved. core/VOCABULARY.md documents these delivered conventions. An ArchUnit rule rejects
Input/Dto/Snapshot/Definition suffixes in canonical model/source/plan packages, with four deliberately
invalid fixture names detected by the negative control.

Before changing snapshots, captured full catalog output from the prior packaged JAR and compared
it with the renamed producer: exact structural equality after only reviewed type/group identifier
substitution (catalog-names-review.json). Counts, field contracts, constraints, defaults, targets
and operations were unchanged. Six packaged discovery outputs were independently compared with
mapped prior snapshots. APPEND_ROW's support-id list has two score-tie segments whose qualified-id
ordering changes with the new names; membership and all other fields stay identical. This is
existing rank/qualified-id ordering, not a search behavior change. Discovery recording then passed
both launchers. names-recorded-snapshot-review.json proves each retained snapshot delta:

- discovery/1-help-protocol.json: one text value, updated model/group example names.
- discovery/2-help-guidance.json: one text value, updated model/group example names.
- discovery/3-help.json: two guidance/flag description values with current group names.
- discovery/10-print-protocol-catalog.json: 135 identifier/reference values.
- discovery/11-print-protocol-catalog.json: 212 identifier/reference values, including the two
  score-tie support-list reorderings described above.
- discovery/12-print-protocol-catalog.json: SET_CELL value-shape reference becomes cellValueSpecTypes.

No execution, recipe, malformed-input or error snapshot changed. Golden normalization is unchanged;
all 141 cases and their value-changing negative control pass in the final full check.

The first focused command also pulled downstream tests through core's aggregate JaCoCo report; use
explicit test/PMD tasks for focused iterations. It stopped on PMD's non-instantiable fixture namespace
rule; converted that fixture namespace to an enum. Focused core/JSON/architecture checks then passed.
Final ./gradlew check passed in 2m47s: core 328, JSON 333, engine 1511, CLI 598, architecture 39;
no failures, errors or skips. No mutation or active fuzzing ran.

Inventories: authored Input files 83 -> 0, Spec 75, Patch eight (a naming migration, not model deletion);
reports remain 96; main Java 124461 -> 124513, tests 154570 -> 154573, converter lines 4695 -> 4700
from wrapping longer names. Engine snapshots 51/definitions 13, Support 217, blank checks 438 remain.
The external inventory tool now reads Union as well as JsonTypeInfo metadata and tracks Spec/Patch
families: it reports actual remaining discriminators, not an empty obsolete annotation search.
Before/after evidence is in wp-23/before and after-names. Baseline diff is empty; supplement absent.

WP-23 remains open: normalize tag fields, unify cell values and reset the wire version to integer 1.
Other pending WPs, including schema emission, remaining identifier validation, central engine
refactoring and final campaigns, remain explicit rather than being counted as done here.

### WP-23 canonical wire tag checkpoint

All 122 tagged production unions now declare type (111) or status (11). WorkbookStep remains the
single structural, untagged envelope. AssertionResult uses status and AssertionStatus; context and
window identities use type and derive their id from Variant metadata rather than repeated literals.
Recipe rows, details and keyword candidates consistently use type/RecipeType. Domain attributes
such as journal verbosity, diagnostic cause stage and named-range backing kind keep their existing
meaning; they are not polymorphic discriminators. No old-name alias or compatibility handler exists.
An architecture rule covers the two tag names and the exact structural exception, with negative
controls for another tag and an unowned empty tag. Three codec regressions start from valid complete
objects, mutate only the tag name and prove removed shape/stage/outcome tags are rejected.

Javac-resolved member migration changed 110 tokens in 37 files, with zero compiler diagnostics
before the rewrite. Captured the prior tag owners, variant ids and record components independently
from the previous packaged artifact (tag-contracts-before.json). Full catalog comparison proves
that only 1254 tag/descriptor fields change; all other fields and values are equal. Current docs
and fixtures were updated. Restored 22 incidental authored-type substitutions in the archived
release changelog from before the naming checkpoint; released historical names remain historical.

Golden review accounts for every change across 56 snapshots: kind -> type (574), view -> type (50),
stage -> type (34), durability level -> type (24), assertion outcome -> status (20), window shape ->
type (two), plus one derived window-group search description saying discriminator type (one).
Total: 705 declared key/description changes. Per-file/path evidence is in
wp-23/tag-snapshot-corrected-review.json. Scalar observations and golden normalization are unchanged.
The review initially missed the host-added protocolVersion envelope on recipe details and falsely
classified a plain named-range surface backing-kind field by its overlapping JSON shape. The full
run detected both. Corrected the review using the producer context: preserve surface.namedRanges
kind and admit the unchanged version envelope. All 46 failed launcher observations then equal the
independently transformed prior snapshots. Retained full failure XML, actuals and difference lists.
One CLI constructor test's expected field label also changed from view to type. The final full run
passes all 141 golden cases and their value-changing negative control.

Final ./gradlew check passed in 3m33s: core 328, JSON 336, engine 1511, CLI 598, architecture 41;
no failures, errors or skips. Named JPMS modules passed 14 requests, catalog, response and twelve
scalars. Inventory: only type/status discriminators; main Java 124513 -> 124474, tests 154573 ->
154674. Spec 75/Patch eight, reports 96, engine snapshots 51/definitions 13, converters 4700,
Support 217, blank checks 438 remain. Registrations, hand descriptors and wire creator/default
sites remain zero. Baseline diff is empty; supplement absent. No mutation or active fuzzing ran.

Challenge for the next cell-value merge: an independently constructed workbook with a stored NaN
at A1 succeeds and reports numberValue="NaN"; a normal C1 control succeeds with 1.0. A NaN formula
cache at B1 currently fails INTERNAL_ERROR. Combined reads fail at that cache, so separate cell
requests were necessary to establish the actual boundary. Source, workbook, request/grant, results
and verdicts are retained in wp-23/nonfinite-observation. Preserve stored numeric observations in
the shared resolved type and keep finite-number checks at the authored expectation boundary. The
formula-cache failure is pre-existing and remains a backend investigation, not a reason to drop
stored observations. WP-23 remains open for cell-value unification and integer protocol version.
Schema defaults, remaining WP-21 inventory targets and all later WPs remain outstanding.

### WP-23 integer protocol header checkpoint

The wire protocol version is now integer 1. Removed the plan enum and its current()/V3 surface;
ProtocolVersion is a validated vocabulary record with one authoritative NUMBER constant. A closed
WireScalar primitive hierarchy extends the existing twelve WireString records with that integer
value. The generic constructor-backed codec still rejects token coercion and preserves owned
field/index paths; it now writes scalar primitive values rather than expanding the record. Graph,
origin and redaction traversal treat the numeric value as a scalar. The catalog unwraps the scalar
component and derives its 1..1 number constraint from FieldRule metadata. No core Jackson dependency
or qualified opens were introduced. Record construction validates version identity once.

All main/test projects compile. Three new regressions independently assert literal integer bytes,
unsupported numeric/string/null/float rejection, rooted problem paths and numeric catalog shape.
The existing twelve-domain string property suite still passes. Generic enum tests were moved from
the retired version enum to the existing assertion-mode enum, preserving allowed-candidate and
empty-enum detection. Constructor-error coordinate assertions now inspect version binding failure
rather than incorrectly expecting an integer to fail enum shape validation. Updated generated test
request builders to emit actual numbers, including the product-based golden readback requests.

The unknown-enum golden fixture now targets execution.assertionMode with a valid integer header;
its code/category and rejection assurance remain intact. Added a separate unsupported numeric
version fixture, with one owned INVALID_REQUEST problem at the exact version property offset.
Corpus: 17 doctor inputs, 72 snapshots and 143 JUnit cases (both launcher paths retained).
Golden review: 115 existing non-doctor header values and 16 doctor response headers change V3 -> 1.
No existing legal payload or scalar workbook observation changes. Independently recomputed 18 doctor
byte offsets and two columns from the edited input bytes; one execution-time missing-input offset
also moves 358 -> 355. The moved enum fixture changes only its diagnostic subject, candidate list,
path and exact property offset. The new numeric-version snapshot proves the new guard. Per-file
reviews are version-snapshot-predicted-review.json, version-doctor-reviewed.json and the retained
execution-error difference/proof. Normalization is unchanged. All 143 cases, including the real
value-changing negative control, pass in the final full run.

Full catalog output equals the prior producer after only the numeric header plus its field's
STRING/enum -> NUMBER/range-1..1 descriptor change (version-catalog-review.json). Named JPMS modules
pass 14 request round trips, catalog, response and twelve string scalars; version transport is also
exercised in every versioned document. No hidden legacy parser or old-string alias remains.

A concurrent authorized task owned shared packaging inputs temporarily; paused shared-file edits
and packaging checks, then verified the aggregate delivered state after it committed. Its committed
legal payloads were used as the snapshot baseline and preserved. All changes in this checkpoint's
index are the protocol work. Recorded intermediate failures include stale enum/type assertions,
a quoted dynamic test-header builder, a source-offset expectation, and a standard PMD literal-order
finding; focused corrections passed before the final full run.

Final ./gradlew check passed in 2m54s: core 328, JSON 339, engine 1511, CLI 600, architecture 41;
2819 tests, zero failures/errors/skips. Inventory: main Java 124474 -> 124528; tests 154674 -> 154714;
JSON main files 69 unchanged (codec files renamed), Spec 75/Patch eight, reports 96, engine snapshots
51/definitions 13, converters 4700, Support 217, blank checks 438 remain. Shell lines 10827 -> 10829.
Creator/default/registration/manual-descriptor inventories remain empty. AGENTS.md baseline diff is
empty; supplement absent. No mutation or active fuzzing campaigns ran.

Next: merge scalar expectations and effective cell observations into one CellValue and remove their
parallel wrapper/converter families. Keep CellReport's meaningful stored-cell variants: they retain
stored type without VALUE projection and enforce type-specific report facets. Rich-run/temporal
metadata belongs to reports, while the shared scalar payload is common with expectations. Preserve
the independently observed stored-NaN behavior and keep authored numeric expectations finite.
Detailed design/challenge notes are external in cell-values-design.md. WP-23 is not complete yet;
WP-21's full identifier/blank target, schema defaults and later phases also remain outstanding.

### WP-23 shared cell-value checkpoint

CellValue is now the common resolved scalar for exact expectations and observed cell values.
Removed CellScalarValue, CellValueReport and the five ExpectedValues wrapper records/conversion
switch. CellReport retains its six meaningful stored-type variants, including FORMULA when VALUE
is unprojected; each projected scalar uses value. Formula runs and temporal metadata now belong
to FormulaReport. Preserved run concatenation, nonempty-run and numeric/temporal invariants, and
preserved the existing requirement that formula metadata needs VALUE while ordinary numeric
TEMPORAL may stand alone. Blank VALUE now explicitly projects the shared blank variant.

Authored CellValueSpec scalar/source/date payloads use value; rich-text run sources use text.
Migrated authored examples, recipes, Java consumers, diagnostic paths and deterministic fuzz seeds.
ExpectedValues factories now return canonical CellValue directly; CellAssertion.Value keeps the
finite-number guard at the authored expectation boundary. Observed NaN/Infinity/-Infinity retain
their encoded string values. New codec tests cover every scalar wire shape, observed nonfinite
round trips and both authored guards. A raw stored-NaN workbook regression exercises the engine,
ordinary numeric control and response round trip. Projection tests independently cover blank
VALUE presence/absence, formula runs/temporal dependencies, and ordinary numeric TEMPORAL alone.

Golden review: 73 initially drifting launcher observations correspond to 37 snapshots. Both launchers
agreed on every one of 36 paired observations; the remaining control snapshot was checked by the
existing real value-changing negative control. After only the declared payload reshapes, all workbook
values/metadata are equal. Formula evaluation scalars move to value and their runs/temporal move
up to the report. Scalar primitives acquire the common typed envelope. Four launcher observations
add the already-implied blank payload. Authored payload keys and rich-run text keys change without
changing their contents. The two old scalar catalog groups merge into cellValueTypes, so compact
index groups drop 72 -> 71 and the cell-text search matches drop 79 -> 78. Search ordering changes
only where group deduplication, renamed fields and truthful variant summaries affect the entries.
Catalog changes are confined to the three affected existing groups, the replacement scalar group,
and action/assertion templates/references. EXPECT_CELL_VALUE's text example now uses the generic
value placeholder "sample-value" instead of the former field-specific "Sample text".

Two diagnostic-coordinate changes were verified against actual serialized fixture bytes: duplicate
stepId offset 238 -> 237 after number -> value; missing-source path offset 355 -> 354 after source
-> value. The missing-source JSON path changes action.value.source.path -> action.value.value.path.
No workbook-root source or chart-title source keys changed. The snapshot normalization code is
unchanged. All 37 recorded snapshots equal the independently retained pre-record actual observations;
the update run passed all 143 cases in 3m30s. Detailed per-case/catalog/offset evidence is external
under wp-23/cell-*. The VALUE facet documentation and user references describe the shared payload.

The first full check exposed eight stale accessors in the separate parity source set. Corrected
those consumers and the focused parity gate passed in 41s, retaining its independent POI checks.
Other earlier failures and their corrections remain in retained logs; no mutation or active fuzz
campaign ran. The full aggregate check and inventory review follow below.

Aggregate ./gradlew check passed in 4m34s: core 328, JSON 343, engine 1513 (including parity),
CLI 600 and architecture/build-logic 41; 2825 tests, zero failures/errors/skips. Named JPMS modules
pass 14 example round trips, catalog, response, twelve string scalars and a shared observed-NaN
report round trip. Deterministic nested Jazzer check passes: 106 support tests and 94 promoted
inputs (48 protocol, 11 workflow, nine engine sequence, 26 XLSX roundtrip). Kept every expected
outcome; only 45 input-byte counts and their text copies needed correction, including the earlier
integer-header shortening. Moved the same rich-run/temporal invariant checks onto formula reports
and validate shared scalar types directly. The compile uncovered this remaining nested consumer;
all three source sets now compile. Retained the stale-count mismatch before correction.

Inventory before -> after: reports 96 -> 95; main Java 124528 -> 124311; test Java 154714 -> 154963;
blank checks 438 -> 437. Spec 75/Patch eight, snapshots 51/definitions 13, Support 217 and JSON files
69 remain. Converter inventory 4700 -> 4701: the common report adapter is still present until WP-40;
its categorization does not capture the removed authoring wrappers/assertion switches. Creator,
wire-default, Jackson-registration and handwritten catalog-descriptor inventories remain empty.
Discriminators remain type/status only. Remaining legacy markers concern stored OOXML formats and
build-directory policy, not wire compatibility handlers. No baseline change; supplement absent.
The shell example materializer and Docker workflow's 20 authored payload fields were migrated too.
Bash syntax passes. Full ./check.sh checkpoint follows this commit; WP-21's aggregate blank target
and schema default publication remain explicitly incomplete, along with the planned engine work.

### WP-23 shell-consumer verification checkpoint

The first full ./check.sh checkpoint passed isolated quality/coverage, nested deterministic replay
and packaging, then failed its shell verifier on stale derived catalog lookup ids. Corrected the
three top-level workbook group ids and executionModeSpecTypes in the verifier and its test doubles.
The adjacent audit found old recipe view tags and V3 string expectations in shell/Python consumers;
updated these to type and integer 1, including discovery-execution fixtures and documentation checks.
Kept independent integer type checks so Python's True == 1 does not accept a boolean header.
No compatibility aliases were added. The artifact verifier and its negative controls pass.

The documentation check also required AFAD front matter on the new ARCHITECTURE.md. D17 explicitly
requires removing AFAD front matter, so removed that obsolete format assertion. Retained request
snippet validation, forbidden inline-secret checks and the deterministic seed inventory assertion.
Its focused test passes. The remaining old group-name audit found no nonhistorical source consumers.
No baseline or supplement changes; no golden changes in this checkpoint.

Full ./check.sh retry passed all six stages in 14m39s, including isolated check/coverage, all 94
promoted replays, packaging, shell regressions, artifact/discovery execution and Docker smoke.
Docker verified legal visibility, signature-line authoring under container fonts, and streaming
write readback. This also supplies passing ./gradlew check evidence for this shell-only commit.
The first failure and focused diagnostic command outputs are retained in wp-23/cell-phase-check*,
cell-shell* and cell-documentation* logs. No mutation or active fuzz campaign ran.

WP-23's discriminator, retired-wire handling and reviewed golden criteria are satisfied. The
aggregate WP-21 blank-check target remains open for the scheduled engine/Jazzer drains, and WP-24
schema publication remains coupled to WP-30. Do not label these incomplete criteria as delivered.
WP-30 starts with the retained unusual-key defect and independent schema generation; grant schema
publication must follow the actual canonical core grant in WP-43 instead of introducing a mirror.

### WP-30 unusual-key intake defect checkpoint

Reproduced the retained S2 failure and adjacent path-loss cases through the packaged CLI. Whitespace
keys could abort syntax recovery with "jsonPath must not be blank"; empty keys lost their JSON path,
and dots/brackets in keys were indistinguishable from path operators. Root "request" also collided
with the root-location sentinel. Added one canonical field-path renderer shared by syntax parsing,
shape findings and token indexing. Ordinary identifiers retain dot paths; unusual keys use JSON-
quoted bracket paths, with the reserved root field quoted separately from the root location.
Token lookup now traverses these canonical prefixes against actual members rather than splitting
arbitrary authored names at dots/brackets. Array indexes retain direct indexed access. Removed the
old ambiguous-key skipping and obsolete field-boundary parsing; first duplicate ownership remains.

The regression first failed seven of eight cases against the old implementation, including the
original blank-path exception. Ten final cases cover empty/whitespace keys, quotes/backslashes,
dots/brackets, the root sentinel, ordinary identifiers, Unicode names, mixed quoted-key/array paths
and UTF-8 byte offsets. Existing empty-key problem-surface tests now require the precise quoted path
instead of absent location. Full JSON tests/static checks pass (353 cases). All 48 deterministic
protocol seeds replay with unchanged expectations; no active fuzz campaign or mutation run occurred.

Only one golden snapshot changes: doctor/blank-key-syntax.bin.json, identically in both launchers.
It previously recorded a generic INVALID_REQUEST blank-path failure. It now preserves nine ordered
findings: four INVALID_JSON grammar findings (unclosed object at offset 0, missing colon at 4,
expected value at 10, missing separator at 23), the unknown quoted whitespace field at 4, and four
missing required root fields (protocolVersion/source/persistence/steps). Verified line/column facts
against the exact UTF-8 bytes; no normalization changes. This intentionally repairs the S2 defect
rather than preserving its collapsed error result. The other 16 doctor inputs are byte-identical.
The full golden suite, including both launchers and the real scalar-changing negative control,
passes in the aggregate run.

./gradlew check passes in 2m51s: 2835 tests, zero failures/errors/skips. Inventory: JSON main files
69 -> 70; main Java 124311 -> 124315; test Java 154963 -> 155058. All model, converter, Support and
blank-check counts are unchanged (95 reports, 51 snapshots, 13 definitions, 4701 converter lines,
217 Support classes, 437 blank checks). AGENTS.md baseline diff is empty; supplement absent.
Detailed red/green tests, before/after packaged observations, golden failure XML and source-location
proof are retained in external wp-30 evidence. WP-30 remains open for schema/intake integration.

### WP-43 canonical grant checkpoint, brought forward for WP-30

WP-30 requires a grant schema, while the canonical core grant appears later as WP-43 (whose only
prerequisite is WP-22). Brought the real grant model forward to satisfy that dependency. No temporary
model or compatibility alias was introduced. ExecutionGrant is now one core record with the existing
read, operation, workbook-scope, publication, secret and acceptance authority. AcceptancePolicy and
its validated requirements moved from engine API to core. Colocated Union/Variant metadata supplies
the existing wire identities. Retained all collection copying/deduplication, operation-id grammar,
selected-sheet validation, exact publication collision policy, nonempty acceptance requirements,
unique assertion/inspection ids and normalized OOXML part-name checks.

Removed the CLI document/authority/policy mirrors, conversion methods, default-grant helper, grant
reader and redundant read-exception wrapper. The template command is part of CliCommand and renders
ExecutionGrant.denied() through json. Missing grants still create no owned scratch space; checked
file I/O now retains the actual failing path and surfaces IO_ERROR instead of a generic internal
failure (before/after packaged observations retained). Grant paths remain rooted at the CLI process
working directory, independently of request-root resolution. Json's immutable per-read context
resolves relative path literals before the core constructor normalizes them. Its Path codec rejects
non-string, blank, null and invalid path inputs and writes filesystem strings instead of URI forms.
A generic mapper read of a relative grant path cannot silently use an ambient directory.

Moved the core acceptance invariant test to its owning module. Replaced the deleted conversion-only
tests with direct core grant codec tests: every authority/policy family, assertion-versus-inspection
rejection, explicit/absent/different roots, lexical path normalization, plain-string output, unknown
fields with otherwise complete input, invalid path scalars and fail-closed round trips. Retained all
engine grant, permission-denial, publication, acceptance and CLI integration tests. Test grant fixtures
now use the canonical model and schedule their owned temporary grant files for deletion on JVM exit.

Focused authority/acceptance/codec/CLI checks and standard static/architecture checks pass. The first
test compilation exposed remaining CLI fixture builders; those now create core grants directly.
The full ./gradlew check passes in 3m46s: core 329, JSON 367, engine 1512, CLI 588, architecture 41;
2837 tests with zero failures/errors/skips. All 143 golden cases, including the actual value-changing
negative control, pass without re-recording any snapshot. Nested Jazzer check passes all 106 support
tests and 94 deterministic seeds in 31s. Named JPMS modules pass the existing 14 requests, catalog,
response, twelve scalars and factual NaN, plus the canonical grant round trip. The packaged engine
JAR contains neither retired grant/policy API class. No mutation or active fuzz campaign ran.

The CLI grant-type inventory is empty (11 -> 0). Main Java 124315 -> 123989, test Java 155058 ->
154955, Support classes 217 -> 216. JSON files 70 -> 72; Spec/Patch 75/eight, reports 95, snapshots
51, definitions 13, converter lines 4701 and blank checks 437 are unchanged. Core remains POI/Jackson
free. AGENTS.md baseline diff is empty; supplement absent. The external metadata prototype derives
250 definitions directly from the real grant and independently validates its serialized authority,
alongside all 14 request examples and two actual responses. Publication of all three schemas and
schema-driven intake remain open in WP-30; do not treat the external prototype as shipped tooling.

### WP-30 schema publication and metadata-truth checkpoint

JsonSchemas now publishes immutable Draft 2020-12 documents for the real WorkbookPlan,
ExecutionGrant and WorkbookResult roots. The explicit :json:generateSchemas task writes the three
committed docs/reference/schema artifacts. Reflection follows canonical components and closed union
variants; properties forbid unknown fields, model presence/default metadata controls requirements,
and FieldConstraint/enum metadata supplies scalar rules. Constant defaults are JSON-native values;
sibling defaults use x-gridgrind-defaultFrom. Request numbers stay JSON numbers; response float/
double observations admit only the existing NaN/Infinity/-Infinity string encodings alongside numbers.
Nonblank patterns derive their whitespace characters from the running JDK. Sorted properties and
metadata-owned variant ids make generation deterministic. The schema comment accurately distinguishes
these structural/declarative checks from domain, token, constructor and workflow invariants.

The new read-only reference check explicitly declares both schema references and examples as Gradle
test inputs. When generation is explicitly requested with tests, mustRunAfter orders the producer;
ordinary check does not regenerate references. A negative control ran the same focused test first
successfully, changed only a reference file, then observed failure without any rewrite. Restored the
exact original bytes before full verification. No hash is being used as a substitute for review.

Ten independent NetworkNT tests validate all 14 examples, a serialized canonical grant (including
paths and a secret-reference object), finite/NaN/+Infinity/-Infinity response values, literal and
sibling defaults and committed references. Wrong-number, string-version, unknown-field, bad grant
path, invalid secret-id and invalid numeric-observation controls are rejected. NetworkNT is pinned
in the version catalog and test-scoped, as selected in S3. Packaged CLI inspection finds neither
NetworkNT nor Jackson 2 databind classes; its transitive validator dependencies do not ship.

A red cross-consumer test reproduced a metadata defect: SecretReference actually serializes as
{id: ...}, but the catalog classified it as STRING. Removed that false special case, added its
colocated plain-type metadata, and shared its unchanged constructor regex with FieldRule metadata.
This changes no accepted wire representation. The full catalog review proves exactly seven fields
change from scalar STRING to the secretReferenceType object reference: SET_SHEET_PROTECTION plus
open/encryption/signature/workbook-protection secret references. The new plain descriptor has one
required id with the existing identifier grammar; all other existing descriptors/templates are equal.
Only discovery/10-print-protocol-catalog.json changes in the golden corpus, identically for both
launchers: one new plain group, 111 -> 112. All other discovery and workbook observations are unchanged.
No normalization changes. Retained the red test, full before/after catalogs and exact difference proof.

Initial iteration failures were two test-tree casts, three standard PMD findings (including two
narrow false positives for immutable-map construction and the generator's operator output), and
stale expected discovery output. Fixed the casts/literal ordering, scoped justified suppressions,
reviewed the sole golden change, and reran focused checks before the aggregate run. No checks were
weakened for data validation. ./gradlew check passes in 3m31s: core 329, JSON 377, engine 1512,
CLI 588, architecture 41; 2847 tests, zero failures/errors/skips, including all 143 golden cases and
the real value-changing negative control. Named JPMS modules pass the earlier model/grant probes
and all three public schema documents. No mutation or active fuzz campaign ran.

Published figures: request 506 definitions / 312304 UTF-8 bytes; grant 250 / 147591; response 552 /
352058. Main Java 123989 -> 124354; tests 154955 -> 155170; JSON files 72 -> 75. Other target counts
remain unchanged: Spec 75, Patch eight, reports 95, engine snapshots 51, definitions 13, converters
4701 lines, Support 216, blank checks 437. Creator/default-site and handwritten-descriptor inventories
remain empty; CLI grant inventory remains empty. Baseline diff is empty; supplement absent.

WP-24's catalog/schema default publication is now delivered alongside its already-empty creator/
wire-default inventories. WP-43's schema prerequisite is shipped. WP-30 remains open: intake still
uses the existing structural validators and Jackson failure normalization. The next work replaces
that path with schema-driven validation and owned domain findings; do not label the new emitter as
completion of the single-intake pipeline or claim standard schemas prove workbook-state preconditions.

### WP-30 emitted-schema structural traversal checkpoint

RequestSchemaStructure now drives root/record/union/collection shape validation from the emitted
request schema, over the duplicate-preserving project tree. Required/unknown fields, type/status
constants, tagged alternatives, structural step alternatives, array items, scalar token kinds and
Java enum candidates come from that document. Java component types are retained only for numeric
fidelity and temporal parsing. Structural step payload names/order derive from the canonical sealed
variants and their schema requirements; no parallel action/assertion/query shape registry remains.

Before switching, compared the new traversal with the old packaged intake for every 17 doctor input
and 14 example: all 31 ordered structural finding lists were equal. An external package-private
probe loaded the old shaded JAR first and the new traversal second, so the comparison remained
independent after switching the working-tree entrypoint. Removed its temporary JUnit comparison once
the public entrypoint used the new path, avoiding a tautological permanent test. Probe source, logs
and the pre-switch characterization source are retained externally.

Deleted RequestNodeValidator, RequestRecordValidator, RequestUnionValidator,
RequestCollectionValidator, RequestNonContainerNodeValidator and RequestNodeTypeClassifier.
RequestStepFragmentCollector now only binds independently valid steps. Removed obsolete record
visibility/optional-field reflection from the raw member index. RequestJavaTypes names the remaining
Java scalar/type-argument contracts; RequestScalarTokens checks numeric and temporal values using
JDK parsing rather than trial Jackson binding. Canonical constructor/domain failures still use the
existing binding path at this checkpoint; declared semantic constraints and their error ownership
must be completed before calling WP-30 done.

Retained primitive-null/required-field inventory checks and retargeted them to the actual schema
traversal. Retained IEEE precision, temporal, valid-sibling/partial-step and scalar wire round trips.
Removed five tests solely covering retired dispatcher/JsonIgnore mechanics and a synthetic optional
carrier already covered through real planId intake. Core models are Jackson-free; CatalogIgnored
remains owned by the schema producer. The public golden assurance is unchanged.

The numeric review found a real unhandled exponent boundary: 1e2147483648 already produced the owned
not-representable result, but 1e2147483649 threw BigDecimal's raw NumberFormatException. Retained that
red failure and now classify the latter as NUMBER_NOT_REPRESENTABLE at the original value token.
The finite/precision-sensitive decimal behavior is otherwise unchanged. All 48 deterministic
protocol seeds retain their expectations. No mutation or active fuzz campaign ran.

After focused JSON/static checks and correcting allocation-only PMD findings, ./gradlew check passes
in 4m25s: core 329, JSON 372, engine 1512, CLI 588, architecture 41; 2842 tests with zero failures/
errors/skips. All 143 golden cases, including the real changed-value control, are byte-identical
without re-recording. Inventories: JSON files 75 -> 71, main Java 124354 -> 124123, test Java 155170 ->
154989, Support 216 -> 215. Blank checks 437 and all model/converter counts remain unchanged.
AGENTS.md diff main is empty; supplement absent.

Remaining WP-30 work: validate constructor/domain invariants before final treeToValue, retain owned
paths/codes without Jackson-message inference, retire failure sanitization and its superseded
helpers/tests, and run the phase gate. Standard schema structural success alone does not establish
constructor/workbook-state validity; the S3 limitation remains explicit.

### WP-30 canonical constructor prevalidation checkpoint

Shape-valid request fragments now pass through their public canonical constructors before the final
Jackson treeToValue. The preflight follows the same closed model graph, decodes Java scalars with
JDK APIs, resolves omission/sibling defaults through ModelDefaults and collects owned invariant
failures. A successful final wire binding must equal the prevalidated model; a later Jackson binding
failure or disagreement is an internal defect, not rewritten into a user-input problem. Tests compare
all 14 examples directly with the final mapper, not with an intake entrypoint that invokes the same
preflight, and independently assert a bad typed step id is rejected before any Jackson binding.

Removed the active request failure normalizer, path guesser and rebaser, their dead code and unused
redactor plumbing. Complete-plan invariant failures use the same owned adapter. Existing coordinate
and typed-exception checks remain; retired tests that bypassed shape validation or asserted message-
inference mechanics were removed/retargeted. Native model constructors now supply formula paths via
the classified formula exception contract, using the existing serializable PayloadLocation and
retaining lower-level causes. Composite assertion emptiness and stored/reported error-literal checks
also own their relative paths. The relative-path adapter appends repeated segments literally; it does
not collapse value.value because of a matching suffix. Focused regressions retain exact formula and
stored-error locations.

The preflight exposed an adopted-vocabulary violation: CustomXmlMappingLocator still used nullable
optional fields. Converted it to Optional<Long>/Optional<String>, updated every caller and removed the
identical engine ExcelCustomXmlMappingLocator mirror plus both identity converters. This is a narrow
prerequisite slice of WP-21/WP-40, preserving the existing wire omissions and mapping behavior instead
of reproducing nullable handling in the new preflight. Core/engine custom-XML tests and integration
checks remain. No compatibility constructors or aliases were added.

Iteration evidence includes the optional-locator constructor gap, retired direct map/scalar binder
tests, missing formula/composite/error paths, exception serialization/cause PMD findings, one stale
engine Optional assertion and the repeated-value path qualifier defect. Focused corrections passed
before the final aggregate retry. The failed aggregate is retained and not concealed by the later
success. No golden snapshots or normalization changed: all 34 doctor observations and all 143 golden
cases, including the real changed-value control, pass on both launchers.

Final ./gradlew check passes in 4m07s: core 329, JSON 372, engine 1512, CLI 588, architecture 41;
2842 tests, zero failures/errors/skips. Named JPMS modules pass 14 request round trips (therefore
canonical reflection without opens), catalog, response, twelve scalars, factual NaN, canonical grant
and three schemas. Refreshed nested Jazzer check passes 106 support tests and 94 deterministic seeds.
No mutation or active fuzz campaign ran. Inventories: JSON files 71 -> 69, converter lines 4701 ->
4687, Support 215 -> 214, blank checks 437 -> 436, main Java 124123 -> 124011, tests 154989 -> 154947.
Spec/Patch 75/eight, reports 95, remaining engine snapshots 51/definitions 13 are unchanged.
AGENTS.md baseline diff is empty; supplement absent.

WP-30 remains open: the shared response/catalog decoder still uses GridGrindJsonProblemMessageSupport
and its framework-error helpers. Remove that mechanism with retained public read behavior, prove the
cleanJacksonMessage/productOwnedJacksonMessage searches empty, and run the full phase gate. Do not
claim completion of the single-intake work while those consumers remain.
