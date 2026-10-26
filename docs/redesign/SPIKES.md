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
