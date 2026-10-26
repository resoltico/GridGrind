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

### WP-30 shared-reader cleanup and phase gate

Replaced the shared codec's framework-prose adapter with JsonReadFailures, which uses typed exception
properties (property name, target class, subtype id, enum value and path references), canonical core
presence contracts and owned constructor causes. It never reads getOriginalMessage, searches wording,
or strips framework suffixes. Removed the old message/value/type/shape/contract detector family,
framework path helpers and duplicate raw step-payload validator: nine production files plus their
sole implementation tests. Retained public response/catalog/doctor IO, stream-ownership, null,
missing-field and round-trip checks. Three direct controls prove arbitrary framework wording cannot
change public classification. Dead synthetic exception/parser fixtures were removed. Standard PMD
findings were corrected by separating typed cause discovery/classification and closing the test
parser; no validation thresholds were changed.

Production Java searches for cleanJacksonMessage, productOwnedJacksonMessage and getOriginalMessage
are empty. One test still calls getOriginalMessage solely to assert text authored by GridGrind's own
step codec; it is verification of our message, not framework-message inference. Valid historical
CHANGELOG_ARCHIVE.md entries and this temporary decision record retain the old names. This is the
explicit scope deviation from WP-30's literal whole-repository grep requirement; E6 permits justified
results, and erasing valid history would not prove removal. No compatibility handler remains.

All 143 golden cases, including both transports and the real value-changing control, remain equal
without re-recording. ./gradlew check passes in 4m22s: core 329, JSON 348, engine 1512, CLI 588,
architecture 41; 2818 tests, zero failures/errors/skips. Reduction from 2842 is retired normalizer/
message-cleaner mechanics, not workbook behavioural checks. NetworkNT independent schema cases,
constructor prevalidation, numeric/temporal and primitive-null protections remain.

The first phase gate passed quality/replay/packaging but failed after 11m08s in a shell string scan:
it confused the response schema's legitimate Admission.operations property with a retired request
field. Corrected that check to parsed authored request envelopes. A real temporary request containing
source + operations fails it, while the generated response schema passes; the control file was
removed. Retained the failed gate and both focused controls. The full ./check.sh retry passed all six
stages in 15m23s, including isolated quality/coverage, deterministic 94-seed replay and support tests,
packaging, shell regressions, real examples/task starters and Docker smoke. Docker additionally
verified odd paths, reopening saved workbooks, signature-line fonts and streaming readback.

Stopped only the recorded task Gradle daemon and its 231 idle compiler workers after the completed
aggregate build, reclaiming roughly 6.8 GiB of retained compiler RSS before isolated verification.
No unrelated daemon/service was stopped. No mutation or active fuzz campaign ran.

Inventory before -> after this cleanup: JSON main files 69 -> 61, Support 214 -> 206, main Java
124011 -> 123345, tests 154947 -> 153624. Converter lines 4687, blank checks 436, Spec/Patch 75/eight,
reports 95, engine snapshots 51/definitions 13 are unchanged. The scoped shell correction adds 20
lines and will be superseded by WP-61's actual Gradle/JUnit product checks. Baseline diff is empty;
supplement absent. Final named-module and inventory outputs are retained under external wp-30.

WP-30 request intake/schema/default publication criteria are delivered, with the historical grep
exception above. Workbook-state preconditions remain in static/runtime checks per S3. Phase 3 is
verified. Next is WP-40's canonical engine model/resolved seam; WP-21's aggregate blank target stays
open until the scheduled engine/fuzz drains. Do not claim the engine, host, tooling, docs or clean-
clone WPs are complete from these intake results.

## WP-40 factual style consolidation, first slice

POI extraction and its consumers now use core CellAlignmentReport, CellProtectionReport,
CellColorReport, CellFillReport, CellGradientFillReport and CellGradientStopReport directly.
Six duplicate engine records/unions and their identity report construction are deleted.
RGB facts retain RgbHex normalization; theme/index and tint retain their exact values.
The authored ExcelColor bridge remains pending its separate core patch migration.

Initial compilation caught the old color package, String/RgbHex and Optional factory consumers;
they were migrated to canonical constructors, including independent saved-workbook and seed
replay readers. Full check initially stopped on standard PMD's import count in the independent
POI fixture. A scoped suppression documents its necessary backend/report correspondence.
Focused engine tests and all 14 parity scenarios passed (31s); nested deterministic check passed
(18s, 106 support tests and every retained input in four replay harnesses). No active fuzzing or
mutation campaign ran; those are deferred until planned implementation is complete per user.

Final ./gradlew check passed in 3m22s: core 329, JSON 348, engine 1498 plus 14 parity, CLI 588,
architecture 14 plus 12 architecture-rule tests; zero failures/errors/skips. All 143 golden cases
passed without snapshot changes. Staged inventories: snapshots 51→45, converter lines 4687→4591,
main Java lines 123345→122963; definitions 13, Support 206, blank checks 436 unchanged. Evidence
is outside the repo under evidence/wp-40. WP-40 remains open: font/border and other mirrors,
resolved plan/step/target and command consumers still need migration. Baseline diff remains empty.

## WP-40 factual font consolidation

Deleted ExcelCellFontSnapshot and ExcelFontHeight. POI extraction, rich-text merging, conditional
formatting, independent readback and replay now consume CellFontReport and FontHeightReport.
The canonical report owns exact fromTwips/fromPoints construction while preserving its two-field
wire shape and mismatched-unit rejection. Five unit/report tests moved from engine to core; the
identity font conversion assertion was retired. Replay unit-consistency arithmetic deliberately
does not call the product factory. Initial compilation caught overbroad constructor substitutions
in the relocated tests; canonical two-argument constructors were restored and all source sets
compiled (22s). No compatibility constructor or engine wrapper remains.

Nested deterministic check passed (28s, all four retained seed harnesses; no active campaign).
./gradlew check passed (3m48s): core 334, JSON 348, engine 1493 plus 14 parity, CLI 588,
architecture 14 plus 12 architecture-rule tests, zero failures/errors/skips. The 143 golden cases
are unchanged. Inventories: snapshots 45→44, converter lines 4591→4566, main Java lines
122963→122914; definitions 13, Support 206 and blank checks 436 unchanged. Older architecture-test
result directories from retired source sets are stale outputs, excluded from these executed-task
counts. Evidence retained outside the repository under evidence/wp-40. Baseline diff empty.

## WP-40 canonical style border and patches

Deleted the identical ExcelCellAlignment/ExcelCellProtection patches and use the core patches
directly, removing both optional identity converters. Deleted ExcelBorderSideSnapshot,
ExcelBorderSnapshot and ExcelCellStyleSnapshot; adapters and readback now return core style
reports directly. Retired the identity whole-style/border-side conversions and their assertions.
No additional core accessor or compatibility factory remains.

Border extraction now chooses the existing None/DefaultColor/Colored alternatives. One old
engine-only test expected an Indexed16 color on a NONE side, which the old public-report converter
already discarded. The test was corrected to assert core None and independently prove the raw
POI indexed color is still retained. A real POI RGB regression proves invisible None and preserved
RGB when switched to THIN. Style merging reads POI directly, so canonicalization does not erase
stored border metadata. No behavioral assurance was replaced by a weaker equality assertion.
Initial compile attempts caught missing canonical test imports and a POI interface/concrete getter;
these were repaired. The obsolete mirror assertion caused the one behavioral failure above; its
full log remains under external evidence/wp-40. Focused style regressions and PMD passed (15s).

Nested deterministic check/replay passed (21s); no active fuzzing/mutation ran. Final ./gradlew
check passed (4m): core 334, JSON 348, engine 1494 plus 14 parity, CLI 588, architecture 14 plus
12 rule tests, zero failures/errors/skips. All 143 golden cases pass unchanged. Inventories:
snapshots 44→41; converter lines 4566→4513; main Java lines 122914→122760. Definitions 13,
Support 206 and blank checks 436 remain open. Baseline diff empty. Next: authored style mirrors,
then remaining factual/model and resolved command consumers; WP-40 remains in progress.

## WP-40 authored color fill gradient and border consolidation

Deleted ExcelColor, ExcelCellFill, ExcelGradientFill, ExcelGradientStop, ExcelBorder and
ExcelBorderSide; adapters, command payloads and deterministic sequence support accept the existing
core specs/patches directly. RGB is unwrapped only for POI string/byte APIs; tint and theme/index
semantics are unchanged. Removed their identity converters and the reverse color identity bridge
in seed support. One valid border requirement test moved to core; two sole-mechanism converter
tests were retired. Real workbook color/fill/gradient/border coverage is retained. No compatibility
constructors/factories were added to core.

Initial compilation caught four remaining String/RgbHex consumers, repaired at their POI/report
boundaries; all engine source sets compiled (25s and 34s). Focused checking first caught an
import-count suppression made unnecessary by converter deletion; removed it and its stale
explanation. Focused engine/core/PMD and all 14 parity scenarios passed (40s). Nested deterministic
check and all retained seed replay passed (22s). No mutation or active fuzz campaign ran.

./gradlew check passed (3m4s): core 335, JSON 348, engine 1491 plus 14 parity, CLI 588,
architecture 14 plus 12 rule tests, zero failures/errors/skips. All 143 golden cases pass without
changes. Inventories: converters 4513→4375, main Java lines 122760→122294; snapshots 41,
definitions 13, Support 206, blank checks 436. Baseline diff empty.

Next font/style migration must preserve cache equality of equivalent authored point/twip heights: the
old converter normalizes to FontHeightReport before MergedCellStyleKey/MergedFontKey equality.
Direct core patch consumption must normalize cache identity at the POI adapter, retain authored
shape elsewhere, and prove value-equivalent style/font reuse with a workbook regression. No extra
engine patch record is justified by that runtime cache requirement.

## WP-40 canonical font and cell-style patches

Deleted ExcelCellFont and ExcelCellStyle and their identity converters, including the reverse font
bridge in deterministic input support. Feature APIs and command payloads now accept CellFontPatch
and CellStylePatch directly. FontHeightSpec exposes exact points/twips queries for POI authoring;
authored units and decimal scale remain intact in the request model. No compatibility factory was
added for retired engine convenience methods. Pure construction/validation tests moved beside
core; sole identity-converter assertions were removed. Initial compile iterations repaired imports,
fully qualified names, generated Optional expressions and target typing.

WorkbookStyleRegistry normalizes only its cache-key patch to TWIPS, preserving the old reuse of
value-equivalent POINTS/TWIPS heights. A real POI test proves identical style identity, only one
additional font/style and unchanged authored POINTS13.0. Its targeted negative control replaced
normalization with identity, failed on the intended duplicate-style assertion in 24s, and restored
production in finally. The normal test and exact-unit tests passed (24s); full focused engine/core,
14 parity and PMD passed (54s). The negative control stayed below the user's 120s interim ceiling;
no full mutation or active fuzz campaign ran.

Nested replay initially caught duplicate protocol/engine style collectors after model unification
and one conditional Optional inference. They are consolidated into the core collector, retaining
unit labels. Its review also fixed the assumption that color-only border patches always have a
style: a direct JDK probe validates color-only plus POINTS coverage labels. Nested deterministic
check and every retained seed passed (19s).

./gradlew check passed (4m43s): core 344, JSON 348, engine 1482 plus 14 parity, CLI 588,
architecture 14 plus 12 rule tests; zero failures/errors/skips. All 143 golden cases are unchanged.
Inventories: converter lines 4375→4343, main Java lines 122294→122046, blank checks 436→434;
snapshots 41, definitions 13 and Support 206 remain open. Baseline diff empty. After the previous
completed gate, this task's daemon 77295 was stopped with 60 idle compiler children (2770MiB
retained); ownership evidence is outside the repo. The next build started a fresh daemon.

Next factual slice review: RichTextRunReport and CommentAnchorReport are exact core counterparts
of their engine snapshots. Rich-text list wrapping, comment conversion methods and raw-cell
projection need producer/consumer migration rather than a substitute wrapper. Comment producers
already filter absent/blank text/author before constructing the old snapshot; preserve that current
observable filtering while removing the duplicated record checks.

## WP-40 canonical rich-text runs and comment anchors

Deleted ExcelRichTextRunSnapshot and ExcelCommentAnchorSnapshot in favor of the identical core
reports. Deleted ExcelRichTextSnapshot: POI extraction returns an immutable List<RichTextRunReport>,
and comments, readback, repair and independent round-trip expectations consume it directly. Removed
run/anchor identity report construction. JDK list-wrapper-only assertions were retired; the actual
POI readback result now has an immutability assertion. Canonical report constructors retain the
nonempty optional-list and concatenation invariants; the raw cell mirror remains pending migration.
No substitute rich-text model wrapper or compatibility API was added.

Compilation caught old runs()/plainText() consumers, which now use lists and JDK joining (19s).
Focused checking first caught the comment repair adapter's import count after crossing packages;
a scoped suppression records its coordinate/anchor/rich-text correspondence. Focused engine,
14 parity and PMD passed (33s); nested deterministic check and all retained seed replay passed
(22s). Root check first stopped on one formatting line after a variable rename (43s); formatting
was corrected. Final ./gradlew check passed (3m36s), core 344, JSON 348, engine 1482 plus 14
parity, CLI 588, architecture 14 plus 12 rule tests, zero failures/errors/skips. All 143 golden
cases are unchanged. No mutation or active fuzz campaign ran.

Inventories: snapshots 41→38, converter lines 4343→4309, main Java lines 122046→121967.
Definitions 13, Support 206 and blank checks 434 remain open. Baseline diff empty. CommentReport
migration still needs moving the removed-model authoring projections to the feature adapter or
source-resolved contract; producer blank-comment filtering is existing observable behavior.

Resolved-plan review evidence: RequestPreflightInputResolver currently reconstructs WorkbookPlan
with inline source-bearing specs; successful PreparedRequest holds that plan and the path-access
bindings. Semantic selectors resolve per step after earlier mutations, so targets cannot all be
made concrete before execution. ExecutionInputBindings.resolveSecret currently grants/checks each
reference, resolves char[] to String and wipes the supplied array; the source-backed text/binary
preflight does not call it. Preserve owned error paths, resource cleanup and secret handling when
introducing actual source-resolved forms and ResolvedPlan/ResolvedStep/ResolvedTarget. Do not wrap
the retired command hierarchy or label remaining authored source references fully resolved.

## WP-40 declared array-formula source gap

Source audit found ArrayFormulaSpec accepts TextSourceSpec (INLINE/UTF8_FILE/STANDARD_INPUT),
but the old source resolver ignored SET_ARRAY_FORMULA and stdin discovery returned false. The
converter required INLINE, so valid file-backed formulas reached INTERNAL_ERROR during mutation.
Three integration regressions failed before the fix (24s): file publication returned INTERNAL_ERROR,
stdin was not discovered, and invalid materialized syntax was classified INTERNAL_ERROR. Full
before-state XML and logs are retained outside the repository.

The existing materialization pipeline now resolves array-formula text, reapplies canonical array
syntax/security validation and collects failures against the original source identity. Stdin discovery
uses the declared source. File/stdin regressions publish and reopen an actual workbook with normalized
SUM(1,2); invalid materialized syntax returns INVALID_REQUEST in RESOLVE_INPUTS at
steps[1].action.formula.source.type and does not publish. The initial test expected the retired
operations root; corrected to the current declared steps root after inspecting the actual example
and WorkbookPlan. All three pass (20s). No schema, requirement or golden snapshot changed.

Focused engine/14 parity/PMD passed (43s), nested deterministic check and every retained seed
passed (23s), ./gradlew check passed (5m16s): core 344, JSON 348, engine 1485 plus 14 parity,
CLI 588, architecture 14 plus 12 rule tests, zero failures/errors/skips. All 143 golden cases pass.
Inventories: main Java 121967→121991, tests 153882→153990; snapshots 38, definitions 13,
converters 4309, Support 206 and blank checks 434 unchanged. Baseline diff empty.

Reflection audit of real record fields finds 16/65 mutation variants and 2/37 selector variants
contain text/binary/secret sources. The other variants can be shared directly. A Java 26 probe
proves source-free records can implement closed authored and resolved contracts together while
only the source-bearing alternative needs a distinct record, with an exhaustive identity switch.
These are design evidence for the pending resolved-plan/step/target work; those WPs are not
claimed complete. No active mutation or fuzz campaign ran.

## WP-40 core materialized cell payloads

Moved the actual source-resolved array formula, rich text, rich-text run and comment contracts to
core as ResolvedArrayFormula/ResolvedRichText/ResolvedRichTextRun/ResolvedComment. These differ
from authored specs because text sources have become strings; source-free font patches remain
shared. Backend APIs consume these core records directly; old engine names are deleted, without
aliases. Deleted the identical engine comment anchor and use CommentAnchorSpec directly, removing
its identity conversion. Pure comment construction tests moved to core.

Compilation caught one qualified nullable type name after relocation (repaired; all source sets
compile in 48s). Focused checks exposed stale capability-inventory source paths, then a core docs
audit that recognized public evidence by the retired contract/ module spelling. Updated that
inventory's contract references to existing core paths, including the resolved-comment row, and
made the audit recognize core without an old-name fallback. A basename collision was resolved
against the model package (not the Java authoring convenience class).

The docs-reading core and parity Test tasks lacked Markdown inputs; declared the top-level docs
and README they read, with relative sensitivity. Both tasks reran and passed (15s). The full
focused engine/core/PMD suites and 14 parity scenarios pass; nested deterministic check and all
retained seed replay pass (20s). Final ./gradlew check passes (3m19s): core 345, JSON 348,
engine 1484 plus 14 parity, CLI 588, architecture 14 plus 12 rule tests, zero failures/errors/skips.
All 143 golden cases are unchanged. No mutation or active fuzz campaign ran.

Inventories: definitions 13→12, converter lines 4309→4297, main Java 121991→121978; snapshots
38, Support 206 and blank checks 434 remain open. Baseline diff empty. These materialized payloads
are prerequisites, not completion of ResolvedPlan/ResolvedStep/ResolvedTarget or command removal.
Drawing review: authored DrawingAnchorSpec permits two-cell only; reports also preserve one-cell
and absolute facts, so a whole-union replacement with the authored type would be incorrect.

## WP-40 shared resolved cell-write contract

Deleted ExcelCellValue and its ten-variant engine hierarchy. ResolvedCellWrite is a closed core
contract that directly shares CellValue's five scalars, the existing authored Date/DateTime records
and ResolvedRichText. Only formula/raw-formula text needs new source-resolved records; their
operand field is value. No duplicate scalar/date/rich-text wrapper or compatibility factory remains.
XSSF, streaming, formula origins, command payloads and deterministic sequence consumers use the
core contract. Source-free dates pass through by identity rather than reconstruction.

Core factual error values include evaluation-only states. Writers retain stored-error validation;
XSSF now validates before clearing an existing formula. A real write-boundary regression proves
that both evaluation-only errors are rejected without erasing that formula. Old factory type-only
assertions were retired; domain validation moved to core, POI behavior stays with the backend.
Compilation caught obsolete expression()/rich-text wrapper accessors (repaired, all source sets
compile in 18s). Standard PMD required a scoped import suppression for the streaming adapter's
canonical operand variants and moving test allocations outside the loop. Focused core/JSON/engine,
14 parity and static checks pass (27s after those repairs). Nested deterministic check and all
retained seed replay pass (17s). No active mutation or fuzz campaign ran.

./gradlew check passes (3m30s): core 346, JSON 348, engine 1483 plus 14 parity, CLI 588,
architecture 14 plus 12 rule tests; zero failures/errors/skips. All 143 golden cases are unchanged.
Inventories: main Java 121978→121920; converter lines 4297→4300 (the remaining authored-to-
materialized conversion is still needed until plan materialization is cut over). Snapshots 38,
definitions 12, Support 206 and blank checks 434 remain open. Baseline diff empty.

This is an actual shared closed-type implementation of the earlier Java probe: unannotated resolved
contracts reuse existing wire records without adding wire variants or a parallel scalar model.
Drawing anchors still need separate authored/report routing before drawing resolved payloads move.

## WP-40 drawing geometry ownership

Deleted the combined engine drawing-anchor union and marker. Writer definitions/APIs now use the
existing DrawingAnchorSpec.TwoCell and DrawingMarkerSpec; readers emit DrawingAnchorReport and
DrawingMarkerReport directly, preserving TWO_CELL/ONE_CELL/ABSOLUTE. Removed report identity
converters and constructor-copy translation. The small remaining twoCellAnchor switch only narrows
the current authored union; it returns the original core record. Obsolete duplicate harness anchor
factories were deleted. No new wire variant or compatibility API was added.

Challenge found a real contract distinction: authoring enforces from/to ordering, while the factual
core report does not impose that authoring constraint. A raw OOXML regression proves reversed
observed coordinates are retained and the authored rectangle still rejects them. Existing one-cell
and absolute observations remain covered. Test fixtures project authored coordinates into expected
report records independently of POI; eight old cross-type equality assertions were updated without
reducing coordinate/behavior coverage. Sole identity-converter assertions were retired.

Initial compilation exposed fixture roles after the split; writer builders now create specs and
read fixtures/oracles create reports (main sources compiled immediately; final source-set compile
9s). Engine/parity/static gate first exposed those eight spec-versus-report equalities, then passed
(29s). A mechanical assertion retarget briefly changed unrelated whitespace; formatting restored
it, and review confirmed only drawing-related test diffs remained. Nested compilation caught six
factual fixture roles, corrected; deterministic check and all retained seed replay pass (21s).
./gradlew check passes (3m44s): core 346, JSON 348, engine 1483 plus 14 parity, CLI 588,
architecture 14 plus 12 rule tests, zero failures/errors/skips. All 143 golden cases are unchanged.
No mutation or active fuzz campaign ran.

Inventories: converter lines 4300→4252, main Java 121920→121771; snapshots 38, definitions 12,
Support 206 and blank checks 434 remain open. Baseline diff empty. Shared source-free drawing
geometry now unblocks moving actual source-resolved picture/chart/shape payloads to core.

## WP-40 resolved chart ownership and shared source-free components

Deleted ExcelChartDefinition. Backend chart preparation/writers consume core ResolvedChart; its
series/plots differ from authored specs because title sources are materialized. ResolvedChartTitle
shares the existing ChartTitleSpec.None and Formula records, adding only the text alternative that
actually changes contract. Axis, legend and data-source mirrors are deleted and their core specs
are reused directly; corresponding identity converters are removed. Shared core chart validation
helpers replace copied list/text helpers. Existing wire annotations stay on the original records;
resolved contracts are not new wire unions. Capability evidence points to the real core record.

All engine/core/parity source sets compiled first pass (33s). Focused core/JSON/engine/14 parity
and static checks passed (70s); nested deterministic check and all retained seed replay passed
(18s). ./gradlew check passed (3m6s): core 346, JSON 348, engine 1483 plus 14 parity, CLI 588,
architecture 14 plus 12 rule tests, zero failures/errors/skips. All 143 golden cases are unchanged.
No mutation or active fuzz campaign ran.

Inventories: definitions 12→11, converter lines 4252→4204, blank checks 434→432, main Java
121771→121635. Snapshots 38 and Support 206 remain open. Baseline diff empty. The shared
closed-type approach now serves real chart titles as well as cell writes, without copying source-free
variants. Remaining resolved drawing/binary inputs and command/target/plan cutover remain WP-40.

## WP-40 immutable binary content and resolved drawing payloads

Moved immutable binary content into core.source.BinaryContent and deleted its engine name.
Preserved defensive construction/access copies, value equality/hash, nonempty authored construction
and explicit empty factual readback. Picture and embedded-object backend payloads now use core
ResolvedPicture/ResolvedEmbeddedObject with actual materialized bytes and shared authored geometry.
No duplicate engine payload or compatibility alias remains. Existing byte-ownership checks moved
from the mixed engine test into core, with caller-buffer isolation added. The first test relocation
attempt found the absent source-test directory; recovered the complete assertions from the preceding
committed source, created the directory and verified the resulting checks.

All source sets compiled first pass (31s). Focused core/engine/14 parity and static checks passed
(52s); nested deterministic check and all retained seed replay passed (18s), including empty
embedded-package preservation. ./gradlew check passed (3m14s): core 347, JSON 348, engine 1483
plus 14 parity, CLI 588, architecture 14 plus 12 rule tests, zero failures/errors/skips. All 143
golden cases are unchanged. No mutation or active fuzz campaign ran.

Inventories: definitions 11→9, converter lines 4204→4203; main Java remains 121635 because these
are necessary materialized contracts moved to their core owner. Snapshots 38, Support 206 and blank
checks 432 remain open. Baseline diff empty. Signature and shape materialization still need their
source-free/optional-field cleanup before the remaining resolved-plan/step/target cutover.

## WP-40 resolved signatures and shapes

Deleted engine signature-line and shape definitions. Core ResolvedSignatureLine carries materialized
preview bytes with Optional<String> metadata; nullable absence is gone from the resolved model.
POI setters use ifPresent, and the old nullable helper is deleted. Shared optional-text normalization
uses the existing signature spec rule. Core ResolvedShape adds only the simple-shape payload whose
text sources have materialized; the existing source-free ShapeSpec.Connector record implements
both contracts and passes through by identity. Geometry remains canonical core data.

Shape name/geometry token rules are shared rather than copied. The interface static methods needed
documentation (PMD caught this in the first focused gate); added explicit responsibility comments.
All source sets compiled first pass (36s, then 32s for the combined shape change). Focused
core/JSON/engine/14 parity/static checks passed (56s), nested deterministic check and all retained
seed replay passed (27s), and ./gradlew check passed (3m): core 347, JSON 348, engine 1483
plus 14 parity, CLI 588, architecture 14 plus 12 rule tests, zero failures/errors/skips. All 143
golden cases are unchanged. No mutation or active fuzz campaign ran.

Inventories: definitions 9→7, converter lines 4203→4201, blank checks 432→428, main Java
121635→121602. Snapshots 38 and Support 206 remain open. Baseline diff empty. Next review
confirmed table columns/styles are source-free and should be reused directly; only table comment
text makes the table payload source-bearing. Do not create resolved column/style mirrors.

## WP-40 typed resolved tables and shared table metadata

Deleted engine table, table-column and table-style model classes. Core ResolvedTable materializes
only comment text; columns and styles reuse TableColumnSpec/TableStyleSpec directly. TableName,
SheetName and RangeRef remain typed in the resolved record, with string unwrapping at POI/workbook
lookup boundaries. Removed column/style identity conversions and the obsolete engine name validator;
remaining selection/pivot references delegate to the canonical TableName constructor. Null-as-empty
record defaults are not carried into the new resolved contract: fixture omissions are explicit.

Initial source compilation exposed only deleted style-converter tests (36s); their identity checks
were retired and named-style validation moved to core. Standard PMD required a scoped allocation
explanation where each name in the still-pending engine selection needs canonical validation. Three
behavioral fixture repairs used typed expected names and explicit empty metadata. Focused
core/JSON/engine/14 parity/static checks passed (38s), deterministic check and all retained seed
replay passed (30s), ./gradlew check passed (4m1s): core 348, JSON 348, engine 1481 plus 14
parity, CLI 588, architecture 14 plus 12 rule tests, zero failures/errors/skips. All 143 golden
cases are unchanged. No mutation or active fuzz campaign ran.

Inventories: definitions 7→5, converter lines 4201→4176, blank checks 428→426, main Java
121602→121534. Snapshots 38 and Support 206 remain open. Baseline diff empty. The engine
selection/query mirrors, remaining factual models and resolved-plan/step/target cutover are still
required; this does not claim WP-40 completion.

## WP-40 resolved XML and validation text; shared validation rules

Deleted engine XML-import and data-validation definitions, prompt/error records and the rule
hierarchy. Core ResolvedCustomXmlImport uses the canonical locator and materialized XML text;
ResolvedDataValidation/ResolvedValidationPrompt/ResolvedValidationErrorAlert carry materialized
text. Source-free DataValidationRuleSpec is reused directly and both rule identity converters are
deleted. Existing public report shapes remain intact. Model bounds/list-copy checks moved to core;
sole type-conversion checks were retired.

Challenge found storage normalization in the deleted engine rule constructors. Preserved trimming
and leading-equals removal at POI write/read boundaries using the existing comparison normalization
owner; optional normalization preserves operand presence so the core comparison contract remains
authoritative. The old non-BETWEEN blank-upper-bound convenience is not retained: core requires
omission, and the obsolete expectation now checks that constraint. XML mapping import accessors
use locator consistently. A scoped import suppression was combined with the XML adapter's existing
annotation after the first repair exposed a nonrepeatable annotation; full failures are retained.

A new real POI boundary regression verifies normalized BETWEEN formulas and that an invalid
replacement formula does not remove the existing validation. Constraint construction now precedes
overlap mutation, preserving the protection formerly supplied by engine model construction. All
source sets compiled first pass (35s); final focused core/JSON/engine/14 parity/static checks pass
(48s), deterministic check and all retained seeds pass (19s). ./gradlew check passes (3m22s):
core 352, JSON 348, engine 1477 plus 14 parity, CLI 588, architecture 14 plus 12 rule tests, zero
failures/errors/skips. All 143 golden cases are unchanged. No active mutation/fuzz campaign ran.

Inventories: definitions 5→3, converter lines 4176→4117, blank checks 426→425, main Java
121534→121422. Snapshots 38 and Support 206 remain open. Baseline diff empty. Remaining
definitions are source-free named-range/pivot/conditional-formatting models; they must be removed
by direct core use, not moved into artificial resolved copies. Print and protection source-bearing
forms and the actual resolved plan/step/target/command cutover remain open.

## WP-40 canonical print settings and materialized text

Deleted engine print layout, setup, margins and header/footer text models. ResolvedPrintLayout
materializes header/footer sources and shares core print-area/scaling/title-band/setup/margin specs.
HeaderFooterTextReport is the identical complete text contract and is reused for materialized text;
blank/empty domain queries do not add wire fields. PrintAreaSpec keeps RangeRef until POI parsing.
Deleted setup/margin snapshots in favor of core reports and removed report identity construction.
The outer raw print snapshot remains pending the full read/query cutover.

Compilation caught a RangeRef label boundary, stale qualified names, and redundant fixture wrapping
(repaired; final source-set compilation 34s). The focused engine suite caught a meaningful diagnostic
difference: the old engine bands rendered Excel-native row/column coordinates, but core constructors
used plain index messages. Preserved that assurance in the canonical core constructors using the
existing ExcelIndexDisplay owner; did not weaken the diagnostic assertions. Core/JSON/engine,
14 parity and static checks then passed (2m16s).

Task daemon 16731 retained 168 idle workers (5447MiB, no active test executors). Ownership/resource
evidence is outside the repo; stopped that task daemon at the completed gate and started fresh.
Nested deterministic check and all retained seeds pass (38s). ./gradlew check passes (4m53s):
core 352, JSON 348, engine 1477 plus 14 parity, CLI 588, architecture 14 plus 12 rule tests, zero
failures/errors/skips. All 143 golden cases are unchanged. No active mutation/fuzz campaign ran.

Inventories: snapshots 38→36, converter lines 4117→4032, blank checks 425→424, main Java
121422→121109. Definitions 3 and Support 206 remain open. Baseline diff empty. Workbook/
sheet secret materialization, the source-free definition drains, remaining factual models and the
actual ResolvedPlan/ResolvedStep/ResolvedTarget command cutover still remain WP-40.

## WP-40 canonical protection facts and resolved credentials

Deleted engine sheet-protection settings and workbook-protection settings/snapshot mirrors.
The adapter now accepts core SheetProtectionSettings and returns WorkbookProtectionReport directly;
the 15-field and five-field identity conversions are gone. ResolvedWorkbookProtection carries
materialized optional credentials with unchanged blank/presence validation and secret binding timing.
Its debug rendering redacts credentials; an independent sentinel regression verifies both redaction
and retained lock state/accessors. Existing real workbook protection roundtrip/clear/hash checks remain.

The first focused run found a now-unnecessary import suppression (58s); removed it. Focused
core/JSON/engine/14 parity/static checks then pass (45s). Nested deterministic check initially stopped
on formatting (6s), corrected with its own Spotless task; check and retained seed replay pass (19s).
Root ./gradlew check passes (3m36s), including all 143 unchanged golden cases; no active fuzz or
mutation campaign ran. Core has 353 passing tests, JSON 348, engine 1477 plus 14 parity.
Inventories: snapshots 36→35, converter lines 4032→3971, main Java 121109→121047; definitions 3,
Support 206 and blank checks 424 remain open. Baseline diff empty. WP-40 remains in progress.

## WP-40 direct authored pivot model

Deleted ExcelPivotTableDefinition and its source/anchor/data-field hierarchy. The actual pivot
adapter accepts PivotTableSpec directly, retaining SheetName/RangeRef/DefinedName/TableName/
CellAddress until the POI boundary. Removed all pivot model conversion methods; source-free
payloads do not get artificial resolved duplicates. Retained workbook lifecycle, health, roundtrip
and POI mapping checks. Moved authored constraints to core tests; explicit empty axis lists isolate
data-field failures and explicit display names replace the deleted engine null/blank convenience.

The first migration script stopped after its mechanical rename because that rename also changed
the converter method name; the initial compile showed six typed scalar boundary errors. Completed
the direct-use edit and scalar unwrapping; all source sets compiled (31s). Focused core/engine,
14 parity and static checks pass (51s); nested deterministic check and all retained seeds pass (18s).
Root ./gradlew check passes (2m58s), all 143 golden cases unchanged, no active mutation/fuzz campaign.
Inventories: definitions 3→2, converter lines 3971→3931, blank checks 424→420, main Java
121047→120840. Snapshots 35 and Support 206 remain open; baseline diff empty. WP-40 remains open.

## WP-40 direct factual comments

Deleted ExcelCommentSnapshot and its authoring projection methods. POI comment producers, cell
metadata, repair readback and sheet results now use CommentReport directly; removed the identity
report conversion. Producer blank-text/author filtering is unchanged. Actual workbook tests assert
text/author/visibility directly; repair assertions also compare exact cell sets and counts. Deleted
constructor/projection-only checks for the retired snapshot; core retains run/text consistency tests.
The deterministic expected-state builder explicitly selects its three plain-comment fields.

All source sets compile (20s). First behavioral/static run had nine PMD test cleanup violations
(33s): unused helper, boolean assertions and loop object creation. Fixed them; engine behavior/static
retry passes (26s), with earlier core/14 parity already passing. Nested replay first found one stale
projection method reference (2s); corrected and deterministic check/all retained seeds pass (17s).
Root ./gradlew check passes (2m57s), all 143 golden cases unchanged. No active mutation/fuzz campaign.
Inventories: snapshots 35→34, converter lines 3931→3919, blank checks 420→418, main Java
120840→120751. Definitions 2 and Support 206 remain open. Baseline diff empty; WP-40 remains open.

## WP-40 shared differential style and border patches

Deleted ExcelDifferentialStyle/ExcelDifferentialBorder and their converters. POI style writing
accepts DifferentialStylePatch/DifferentialBorderPatch directly; optional sides are unwrapped only
at the OOXML boundary. Font heights remain authored specs for writes and factual reports for reads;
sheet-copy materialization uses exact twips. Rule styles pass directly without identity maps.
Retained POI font/color/border and copy regressions; removed converter-only assertions.

Compilation exposed seven production boundary mismatches, then 23 and four stale fixture/converter
uses; repaired actual optional/scalar boundaries and fixtures. Focused core/engine/14 parity/static
checks pass (41s). Nested compile found one stale qualified fixture name (8s); corrected. Deterministic
check/all retained seeds pass (21s). ./gradlew check passes (2m59s), all 143 golden cases unchanged.
No active mutation/fuzz campaign ran. Inventories: converter lines 3919→3874, blank checks 418→417,
main Java 120751→120662; snapshots 34, definitions 2 and Support 206 remain open. Baseline diff empty.

## WP-40 core conditional-formatting rules and factual sheet copy

Deleted ExcelConditionalFormattingRule and ExcelConditionalFormattingThreshold. Adapters accept
existing core rule/threshold variants directly; removed their conversion switches. Record-order
fixture edits preserve the authored meaning. Authored formulas remain unchanged in core and normalize
at POI creation. Two real POI regressions failed before correction (12s; full XML retained): padded
formula storage and invalid replacement preservation. They now pass; rule creation/validation precedes
removing overlaps. Removed conversion-only tests and used explicit barriers for styleless authored
fixtures. Factual styleless rules remain supported through the copy path below.

Challenge: parity's advanced sheet-copy case failed because existing styleless, non-barrier rules
cannot pass strict authored constructors. The former engine mirror had masked that distinction.
Corrected the method, not the authored requirement: snapshot validated conditional-format XML, keep
unsupported-rule/style rejection before cloning, copy XML and retarget formula operands and formula
thresholds directly. This preserves observed rules without authored reconstruction and deletes the
report-to-authoring conversion chain. Runtime copy snapshot holds CTConditionalFormatting payloads.
New real workbook regressions verify styleless formula retargeting/source preservation and unsupported
style rejection before destination creation. No new compatibility model or authored variant added.

Initial compile caught the factual UNALLOCATED enum case and stale conversion-only assertions.
Parity failure drove the copy correction; subsequent compile/static failures were missing import,
nullable style annotation and a loop allocation. Focused engine/14 parity/static checks pass (25s);
all new regressions pass. Nested deterministic check/all retained seeds pass (16s). Root ./gradlew
check passes (3m02s): core 355, JSON 348, engine 1476 plus 14 parity, CLI 588, zero failures/errors/skips.
All 143 golden cases unchanged; no active mutation/fuzz campaign. Inventories: converter lines
3874→3779, blank checks 417→416, main Java 120662→120273. Snapshots 34, definitions 2 and Support
206 remain open. Baseline diff empty. The selected-range block and named-range definition, factual
model drain and real resolved plan/step/target execution cutover remain WP-40 work.

## WP-40 direct print-layout reports

Deleted ExcelPrintLayoutSnapshot. The POI adapter returns PrintLayoutReport with sheet identity,
print-area/scaling/title-band facts, header/footer and setup. Inspection passes that report directly;
removed the runtime print conversion and its now-unused leaf conversion helpers. Existing resolved
print state is still used by sheet-copy materialization; this checkpoint does not claim that separate
copy/read path has been consolidated. Removed one unused fixture helper and one duplicate setup read.

All source sets compile (20s); engine behavior/14 parity/static checks pass (40s). Deterministic check
and all retained seeds pass (16s). ./gradlew check passes (3m03s), all 143 golden cases unchanged.
No active mutation/fuzz campaign. Inventories: snapshots 34→33, converter lines 3779→3723, main Java
120273→120241; definitions 2, Support 206 and blank checks 416 remain open. Baseline diff empty.

## WP-40 shared hyperlink variants

Deleted ExcelHyperlink and its duplicate type enum. Engine commands, POI writers/readers, metadata
and deterministic expected-state fixtures use HyperlinkTarget directly. Removed request/report
identity conversions; POI type/address switches handle the existing four variants at the adapter.
Email uses its canonical email accessor and file uses path. Authored normalization/validation checks
moved into core; retired constructor/type-label and conversion-only checks. POI hyperlink behavior,
relative file resolution, readback and analysis checks remain. Raw hyperlink health classification
retains its existing nonthrowing checks. The canonical core URL allowlist governs authored values;
the raw health analysis still inspects workbook addresses independently.

Production compilation passed; 16 stale test accessor/converter references were migrated. Focused
core/engine/14 parity/static checks pass (45s); deterministic check/all retained seeds pass (17s).
Root ./gradlew check passes (3m09s), all 143 golden cases unchanged. No active mutation/fuzz campaign.
Inventories: converter lines 3723→3690, blank checks 416→415, main Java 120241→120109; snapshots 33,
definitions 2 and Support 206 remain open. Baseline diff empty. WP-40 remains in progress.

## WP-40 direct factual pivot reports

Deleted ExcelPivotTableSnapshot and its source/anchor/field/data-field hierarchy. The pivot reader
constructs PivotTableReport directly, retaining factual range labels and existing typed identities.
Removed the runtime pivot conversion and passed report lists directly. Runtime PivotLocation and
column-axis structures remain backend concepts, not wire mirrors. Updated scalar analysis evidence
and real workbook assertions to read typed values; retained save/reopen, lifecycle and health checks.

Initial compilation exposed one RangeLabel evidence boundary; all source sets then compiled (23s).
Behavior/static/parity run passed except three raw-string-versus-value-object assertions (39s).
Focused retry exposed two further typed source fields (12s); corrected and focused tests pass (15s).
Nested compile caught scalar shape assertions for typed report fields (2s); corrected, deterministic
check/all retained seeds pass (20s). ./gradlew check passes (4m13s), all 143 golden cases unchanged.
No active mutation/fuzz campaign. Inventories: snapshots 33→32, converter lines 3690→3613, blank
checks 415→413, main Java 120109→119901; definitions 2 and Support 206 remain open. Baseline diff empty.

## WP-40 factual conditional-formatting reports

Deleted engine conditional-formatting block/rule/threshold and differential-style snapshots.
POI reads now construct ConditionalFormattingEntryReport/RuleReport/ThresholdReport and
DifferentialStyleReport directly. Color and border snapshot internals use CellColorReport and
DifferentialBorderReport/SideReport, preserving observed metadata without authoring projections.
Removed all corresponding runtime conversions and the newly unused InspectionResultCellStyleReportSupport.
Copy validation consumes optional factual styles. Health analysis retains malformed-range checks,
formula checks and evidence after unwrapping typed labels/options at the appropriate boundaries.

Moved core report invariants into core tests and retired conversion-only assertions. Real POI,
workbook/copy, malformed-style, health and parity checks remain; a test-only DifferentialStyleFixtures
maps independently authored fixtures to expected factual colors/borders without sharing production
read logic. Compilation exposed range/optional boundaries, stale report expectations and test-fixture
annotations not on that source set; repaired without a new dependency. The first full behavior pass
had ten typed-wrapper assertion failures (45s); two and one further sparse-border assertions surfaced
on retries. Final focused malformed-style checks pass (15s). Deterministic check/all retained seeds
pass (23s), then root ./gradlew check passed (3m10s), golden unchanged.

Added a real POI theme-index-99 readback regression. My initial test incorrectly assumed ColorSpec
rejects 99; the observed report assertion passed, but that rejection assertion failed. Verified both
contracts allow nonnegative indices, removed the invalid assertion and retained its failure XML.
Final ./gradlew check passes (2m45s): core 363, JSON 348, engine 1463 plus 14 parity, CLI 588, zero
failures/errors/skips; all 143 golden cases unchanged. No active mutation/fuzz campaign ran.
Task daemon 3833 had 125 retained workers at an idle completed gate; recorded ownership/resource
state and stopped that task daemon only. Fresh checks ran afterward.

Inventories: snapshots 32→28, converter lines 3613→3442, Support 206→205, blank checks 413→410,
main Java 119901→119504. Definitions 2 remain; baseline diff empty. WP-40 remains open.

## WP-40 direct factual table reports

Deleted ExcelTableSnapshot/ExcelTableColumnSnapshot/ExcelTableStyleSnapshot. POI table catalog,
analysis, copy materialization and semantic selectors use TableEntryReport/ColumnReport/StyleReport.
The runtime table conversion and leaf conversions are gone. Labels remain typed until parsing,
lookup, POI or string evidence boundaries. Optional column metadata remains optional for readback;
copy materialization supplies explicit empty authored fields where absent. Compiler incompatible-type
checks caught and corrected label/string comparisons in document analysis and table routing.

Initial compile exposed scalar/optional boundaries; follow-up compile caught four equivalent test
collection/filter comparisons. Removed one now-unnecessary import suppression (24s). Behavior/parity/
static run had fourteen stale wrapper assertions (45s), then one introspection assertion (25s).
The first targeted edit missed that exact expression (3s); corrected it, focused test passes (11s).
Deterministic check/all retained seeds pass (15s). ./gradlew check passes (3m18s), all 143 golden
cases unchanged, no active mutation/fuzz campaign. Inventories: snapshots 28→25, converter lines
3442→3374, blank checks 410→409, main Java 119504→119283; definitions 2 and Support 205 remain.
Baseline diff empty. WP-40 remains open.

Adjacent autofilter review found a pre-existing codec contradiction: Dynamic criteria in core spec
and report use component `type` as domain payload while the union also owns `type` as discriminator.
A packaged-JAR codec probe of Dynamic("today",...) throws InvalidDefinitionException for that exact
property collision (external evidence dynamic-filter-wire-probe.log). The next autofilter slice will
use a distinct domain field, with codec regressions; this follows D5's discriminator ownership and
the user's delegated QA authority without adding aliases. Table work above is independent and green.

## WP-40 dynamic-filter discriminator collision correction

Confirmed the adjacent contradiction with the packaged codec: serializing Dynamic("today",...)
throws InvalidDefinitionException because the domain field `type` collides with union discriminator
`type`. Both authored and factual core criteria were affected; generated schemas even had duplicate
`type` required entries. Renamed the domain component to `dynamicType`, retaining `type: DYNAMIC` as
the sole variant discriminator. No aliases or old-field handling added. Updated consumers, diagnostics,
mutation guide and explicitly regenerated request/response schemas. Three codec regressions cover
authored roundtrip, factual roundtrip and schema-driven request intake with the distinct domain field.
Adjacent plain `type` attributes on non-union records do not have this collision.

Focused codec/schema generation passes (18s); deterministic check/all retained seeds pass (16s).
Root check first found one expected diagnostic still naming `type` (37s); corrected the expectation,
focused core retry passes (5s). Final ./gradlew check passes (2m53s), all 143 existing golden cases
unchanged; no snapshots re-recorded. No active mutation/fuzz campaign. Inventory counts unchanged:
snapshots 25, definitions 2, converter lines 3374, Support 205, blank checks 409, main Java 119283.
Baseline diff empty. This corrects the verified codec defect consistently with D5 discriminator
ownership under the user's delegated QA authority; autofilter model consolidation follows.

## WP-40 direct autofilter specs and reports; factual range labels

Deleted four authored autofilter mirrors (column/criterion/sort-state/sort-condition) and five factual
snapshot models. POI adapters accept existing core specs and produce existing core reports directly.
Removed request/report conversions; copy materialization remains an explicit report-to-authored
boundary. Sort ranges unwrap only at OOXML/parser boundaries, dynamic optional numeric values use
ifPresent, and fuzz generators no longer name deleted Excel models. Compiles found six production
boundaries, one obsolete converter-only assertion and nine unnecessary nullable wrappers around
primitive fixtures; corrected those and retained behavior/health/copy tests.

Challenge: the existing real workbook test preserves a blank stored autofilter reference, while
RangeLabel rejected blank text despite its factual role. Corrected RangeLabel to retain all nonnull
stored text, including blank/malformed references; authored RangeRef remains strict. Added core
regressions for both contracts, extended JSON scalar roundtrips to blank labels (retaining null and
non-string rejection), updated vocabulary, and corrected only obsolete factual blank-rejection
assertions. Core/JSON range-authoring rejection checks remain. Initial engine run had nine stale
scalar assertions including this genuine mismatch (26s). Subsequent runs surfaced the old generated
invalid-label assumption and three other blank-report expectations, then one stream-built label
expectation (28s); corrected those against the observed workbook contract. Final focused test passes
(11s); deterministic check/all retained seeds pass (15s).

Final ./gradlew check passes (2m56s): core 365, JSON 351, engine 1463 plus 14 parity, CLI 588, zero
failures/errors/skips; all 143 golden cases unchanged. No active mutation/fuzz campaign. Inventories:
snapshots 25→20, converter lines 3374→3188, blank checks 409→396, main Java 119283→118644; definitions
2 and Support 205 remain. Baseline diff empty. WP-40's remaining factual models, source-free
named-range/block definitions and real resolved plan/step/target execution cutover remain open.

## WP-40 canonical named-range authoring and observed scope

Deleted ExcelNamedRangeDefinition/Target/Scope/Snapshot. Named-range authoring accepts the existing
source-free StructuredMutationAction.SetNamedRange directly; coordinate spelling remains authored
until the POI formatter creates absolute normalized references. Core NamedRangeReport flows directly
through reads. NamedRangeScopeReport is a distinct observed contract with SheetLabel, while authored
NamedRangeScope retains SheetName validation. Updated analysis/surface scope consumers and removed
identity conversions. Formula-backed observations retain empty stored references; range-backed reports
still require a nonblank reference. Real baseline probes passed before migration (12s) for empty
formulas and malformed stored local sheet labels. Both probes and a new codec roundtrip pass afterward.

Sheet copy uses observed reports directly instead of reconstructing authored definitions; removed
its unused test-only range filter and no-op guard. Unicode/name validation moved to core. New POI
parameterized tests verify four absolute range encodings, including reversed endpoints, without
changing authored coordinates. Analysis evidence uses explicit workbook/sheet domain labels rather
than record implementation names. Existing copy, lookup, collision, source and parity behavior remains.

Compilation exposed typed identifier/scope boundaries, one mechanical helper-prefix collision and
obsolete converter assertions; repaired them. Catalog initialization required declaring the referenced
report scope as a catalog union; fixed metadata, not catalog fallbacks. PMD caught loop allocations,
resolved by sharing the actual range formatter and validating destination identity once. A malformed
test edit was caught by formatting (1s) and corrected. Engine assertion retries corrected six typed
values and one old record-debug evidence string. Core/JSON/engine/14 parity/static checks pass (39s).
Nested compile found one string-key boundary and two report-scope fixtures; corrected, deterministic
check/all retained seeds pass (13s). The first root run then exposed only the expected catalog golden
drift in both launchers (3m08s).

Golden review: exactly one file changed, discovery/10-print-protocol-catalog.json, by three lines.
It adds nested group namedRangeScopeReportTypes with WORKBOOK/SHEET, necessary to describe the
verified distinct factual scope contract. No authored scope group or workbook observation changed.
Reviewed semantic diff is external evidence; discovery factory re-record/replay passes (6s). This is
a narrow metadata deviation from WP-40's unchanged-corpus wording under delegated QA authority.
Final no-update ./gradlew check passes (3m): core 374, JSON 352, engine 1456 plus 14 parity, CLI 588,
zero failures/errors/skips; all 143 golden cases pass. No active mutation/fuzz campaign ran.

Inventories: snapshots 20→19, definitions 2→1, converter lines 3188→3124, blank checks 396→391,
main Java 118644→118471; reports 95→96 for the distinct observed scope contract. Support 205 remains.
Baseline diff empty. WP-40 remains open: factual models, selected-range block, resolved execution seam.

## WP-40 direct package-security and array-formula facts

Deleted engine OOXML encryption/signature/package-security and array-formula snapshots. POI readers,
open/admission context and workbook inspection now share core reports; removed runtime conversions.
Core package-security report owns plain-state construction and the encryption/signature presence
query. Signature invalidation remains in persistence support, where the mutation decision is owned;
it maps loaded signature reports through the existing signature-state transition. Retained the real
signed-workbook regression that reads VALID, mutates through execution, reads INVALIDATED_BY_MUTATION,
and verifies persistence behavior. Removed only retired model-transition assertions.

Compilation caught those two deleted transition calls and two typed array-range comparisons.
Behavior/static/parity run then exposed address/sheet assertions and a natural sort of RangeLabel;
corrected tests to compare/sort their actual scalar values. Focused array tests pass (11s), retained
security behavior/static/parity checks pass, deterministic check/all retained seeds pass (15s).
Final ./gradlew check passes (3m), all 143 golden cases unchanged; no active mutation/fuzz campaign.
Inventories: snapshots 19→15, converter lines 3124→3067, blank checks 391→389, main Java
118471→118271; definition 1 and Support 205 remain. Baseline diff empty. WP-40 remains open.

## WP-40 direct custom-XML reports

Deleted custom-XML mapping/schema/data-binding/linked-cell/linked-table/export snapshots and the
settings mirror. POI introspection returns existing core reports directly, using their established
nullable optional fields and typed observed coordinates. Removed the entire structure-report
converter; mapping and export inspection pass through canonical reports. Kept XML import/export,
malformed metadata, linked table/cell and schema behavior checks.

Compilation exposed two ternary optional-unwrapping mistakes in producer edits; rewrote those
expressions clearly. Test compile exposed old optional schema accessors, a missing assertion import
and one redundant optional wrapper. Behavior/static/parity checks surfaced six representation/fixture
assertions, including a null fixture that failed before its intended invalid-map-ID check; corrected
fixture construction so the actual domain failure is exercised. Final targeted XML tests pass (12s),
retained core/engine/14 parity/static checks pass, deterministic check/all retained seeds pass (15s).
Final ./gradlew check passes (3m27s), all 143 golden cases unchanged; no active mutation/fuzz campaign.
Inventories: snapshots 15→9, converter lines 3067→2976, Support 205→204, blank checks 389→383,
main Java 118271→117984; definition 1 remains. Baseline diff empty. WP-40 remains open.

## WP-40 direct sheet-presentation specs and reports

Deleted engine presentation/display/outline/default/ignored-error mirrors and the presentation
snapshot. Writes accept existing SheetPresentationSpec and leaf specs; reads construct factual
presentation/leaf/ignored-error reports. Removed the request presentation converter and report
identity conversion. Deleted the snapshot's unused authoring-projection API and tests for that API;
pure model validation moved into core, while POI enum mapping and workbook behavior remain engine tests.

Preserved ignored-error range normalization at the POI boundary using the existing parsed-range
formatter. A new real POI regression verifies authored B2:A1 remains unchanged and reads back A1:B2.
Compilation caught report imports, one mechanical prefix artifact, old projection calls and mixed
spec/report fixture lists. Behavior/static/parity tests found four spec-versus-report comparisons;
corrected independent factual expectations. One expected defaults pair was initially copied from the
wrong fixture, corrected to its actual authored values. Final core/engine/14 parity/static checks
pass (26s), deterministic check/all retained seeds pass (17s). ./gradlew check passes (2m56s), all 143
golden cases unchanged, no active mutation/fuzz campaign. Inventories: snapshots 9→8, converter
lines 2976→2909, main Java 117984→117784; definition 1, Support 204 and blank checks 383 remain.
Baseline diff empty. WP-40 remains open.

## WP-40 factual validation reports without authored source wrappers

Deleted the validation snapshot and the entire validation report converter. POI reads construct
DataValidationEntryReport and its factual definition directly. Added distinct prompt/error reports
with observed string title/text; authored source specs and resolved write payloads remain separate.
The former converter fabricated INLINE sources from observed text, contradicting D1's distinct
contracts. Removed that projection rather than moving it. Preserved rule normalization, prompt/error
presence filtering, selected clear/split behavior, overlap analysis and malformed-rule diagnostics.
Health evidence unwraps range labels only at parser/evidence boundaries. Documented the response
shape and explicitly regenerated schemas. A codec regression checks observed message roundtrips.

Readback fixtures independently map authored values to expected reports without sharing production
read logic. Report invariants moved into core; obsolete converter-only tests retired. Compilation
found range/evidence boundaries, resolved-versus-factual fixture definitions and deleted converter
calls. Behavior retries corrected ten and then four stale range/metadata assertions, including one
obsolete factual blank-label rejection. Final core/JSON/engine/14 parity/static checks pass (31s),
deterministic check/all retained seeds pass (17s). The first full root run then found exactly six
expected golden mismatches in both launchers (3m03s).

Golden review/re-record: three files only. Catalog snapshot 10 adds the two factual report groups;
search snapshot 11 adds exactly those two matches and totalCount 78→80, preserving all existing
matches/order. DATA_ENTRY_WORKFLOW changes exactly eight observed leaves (four in execution
inspections, four in reopened workbook facts): prompt/error title/text INLINE objects become the
same strings. A structural comparison independently verifies every other value, key, list length,
authored source object, range, flag, rule and message text is unchanged. Review JSON retained outside
the repo. Normal discovery/example/recipe update replay passes (2m05s). This is a narrow D1 contract
correction under delegated QA authority, explicitly deviating from WP-40's unchanged-corpus wording.

Final no-update ./gradlew check passes (2m41s), all 143 golden cases pass. No active mutation/fuzz
campaign. Task daemon 56897 retained 123 workers/6056MiB at the completed idle gate; recorded evidence
and stopped that task daemon only. Inventories: snapshots 8→7, converter lines 2909→2843, Support
204→203, blank checks 383→382, main Java 117784→117727; reports 96→98 for distinct factual text
contracts. Definition 1 remains. Baseline diff empty; WP-40 remains open.

## WP-40 direct chart reports

Deleted ExcelChartSnapshot and its title/legend/axis/series/data-source/plot hierarchy. Chart reads
and drawing inventory now use ChartReport directly; removed the whole chart identity-conversion
chain. Old three-field series fixtures now supply their four explicit optional presentation fields.
Retired the standalone converter-switch test; retained real POI plot-family/cache/title behavior and
inspection-result coverage. Mixed chart/XML inspection coverage now exercises the actual inspection
result path rather than a deleted leaf converter.

Compilation caught one helper import changed by a mechanical prefix match and four deleted-converter
references; corrected those. Core/engine/14 parity/static checks pass (53s), deterministic check/all
retained seeds pass (17s). ./gradlew check passes (2m53s), all 143 golden cases unchanged, no active
mutation/fuzz campaign. Inventories: snapshots 7→6, converter lines 2843→2640, blank checks 382→379,
main Java 117727→117154; definition 1 and Support 203 remain. Baseline diff empty. WP-40 remains open.

## WP-40 direct drawing and signature-line reports

Deleted drawing-object and signature-line snapshots, including their signature setup/preview mirrors.
POI drawing and VML signature readers return DrawingObjectReport and its nested contracts directly;
removed the runtime drawing conversion and signature inventory identity construction. The actual
producer still reports its stored anchor shape; no authored geometry substitution or compatibility
model was introduced. Binary payload extraction remains a separate next step.

All source sets compile without repair. Core/engine/14 parity/static checks pass (42s), deterministic
check/all retained seeds pass (16s). ./gradlew check passes (3m21s), all 143 golden cases unchanged.
No active mutation/fuzz campaign. Inventories: snapshots 6→4, converter lines 2640→2565, blank checks
379→369, main Java 117154→116807; definition 1 and Support 203 remain. Remaining snapshots are the
genuine HostAcceptanceSnapshot/ColumnAxisSnapshot runtime concepts and the cell/metadata mirrors.
Baseline diff empty. WP-40 remains open.

## WP-40 canonical drawing payload reports

Deleted ExcelDrawingObjectPayload and the remaining drawing report converter. POI extraction
constructs DrawingObjectPayloadReport directly and encodes bytes at that read boundary; runtime
inspection passes reports through. Internal binary readback/copy objects retain their actual byte
contracts. Real picture/raw/POIFS/embedded/copy/empty-payload assertions now decode report base64
before comparing exact bytes. Removed only the retired leaf-converter assertion.

Production compiles; test compile found four copied-payload accessors and the old converter call.
Corrected those without changing internal byte-only tests. Core/engine/14 parity/static checks pass
(50s), deterministic check/all retained seeds pass (15s). ./gradlew check passes (3m28s), all 143
golden cases unchanged; no active mutation/fuzz campaign. Inventories: converter lines 2565→2529,
Support 203→202, blank checks 369→364, main Java 116807→116687; snapshots 4 and definition 1 remain.
Baseline diff empty. WP-40 remains open; cell/metadata mirrors and the resolved execution seam remain.

## WP-40 canonical cell reports at the POI read boundary

Deleted ExcelCellSnapshot/metadata and duplicate engine projection/facet types. POI cell reads now
produce CellReport with shared CellValue scalars directly; removed the runtime cell converter and
metadata grouping. Public commands pass canonical projections unchanged. Internal no-projection
cell/window helpers explicitly request full facts; preview requests VALUE/FORMAT, semantic keys
request VALUE/FORMULA, and schema inference requests the facts it actually uses. Formula results
are scalar values, not recursively nested cell snapshots. Date/temporal interpretation happens at
the POI boundary with the workbook date system. Runtime inspection no longer reconstructs cell models.

Replaced retired snapshot/converter/reflection-only tests with real workbook facet tests covering
formula/value/format masks, missing-cell projections, rich text, metadata, temporal/schema inference
and formula-derived temporal values. Saved metadata fixtures use independent typed hyperlink/comment
accessors. Retained real sheet style/copy/selection/error/preview and public executor behavior checks;
fixture construction uses canonical report fields rather than a replacement raw snapshot type.
Compilation and behavior retries corrected optional/typed accessors and several overbroad test edits,
including unrelated rich-text list access and default preview formatting. Final core/JSON/engine,
14 parity/static checks pass (42s); deterministic check/all retained seeds pass (17s).

Resolved the pre-existing nonfinite formula observation failure retained from WP-23. A real public
executor regression now includes a stored NaN numeric cell, a normal numeric cell and a formula with
stored NaN. Direct diagnosis showed unrequested DataFormatter execution mutating/erroring the formula
and then throwing a numeric/error mismatch. VALUE-only reads now skip formatting and preserve the
actual evaluator NaN result. An initial diagnostic used an inaccessible test helper (compile failure),
then the public rooted temp factory exposed the full cause. My first guard edit missed formatted
source, and an intermediate expected-error assumption was disproved by the returned numeric NaN;
corrected both against direct output. The regression passes and roundtrips the complete response.
This establishes VALUE-only nonfinite observation, not all possible FORMAT/temporal combinations.

Final ./gradlew check passes (2m57s): core 377, JSON 353, engine 1447 plus 14 parity, CLI 588, zero
failures/errors/skips; all 143 golden cases unchanged. No active mutation/fuzz campaign. Inventories:
snapshots 4→2, converter lines 2529→2289, Support 202→201, main Java 116687→116301; definition 1 and
blank checks 364 remain. Remaining HostAcceptanceSnapshot and ColumnAxisSnapshot are actual runtime
concepts (host acceptance evidence and parsed pivot axis state), not core model mirrors. Baseline
diff empty. WP-40 remains open: selected-range block, resolved plan/step/target and command/query cutover.

## WP-40 conditional-formatting targets without a model block

Deleted the last engine Definition model, ExcelConditionalFormattingBlockDefinition. The adapter
accepts canonical RangeRef targets and ConditionalFormattingSpec independently; retained empty and
duplicate range validation at the actual POI mutation boundary. A real workbook regression confirms
both rejected range-list forms preserve existing rules. Removed constructor/converter-only checks;
retained actual rule persistence, copy, command execution and inspection checks. The deterministic
sequence generator now constructs the command directly, retaining range-first byte consumption.
Its successful-value sweep explicitly selects valid ranges because canonical RangeRef rejects the
intentionally malformed alternatives earlier; real harnesses retain expected-invalid handling.

Focused core/engine/parity/PMD checks pass (43s). Nested compilation first found three stale test
fixtures, then the value sweep exposed earlier canonical range rejection; both were corrected and
failure output retained externally. Final nested check/all retained seeds pass (11s); no active
fuzz or mutation campaign. Final ./gradlew check passes (2m57s): core377, JSON353, engine1447 plus
14 parity, CLI588, zero failures/errors/skips. All143 golden cases unchanged. Inventories:
definitions1→0, snapshots2 unchanged (runtime concepts justified above), converter lines2289→2283,
main Java116301→116274, blank364 and Support201 unchanged. Baseline diff empty. WP-40 remains open
for resolved execution and removal of command/query/selector translation.

## WP-40 executable core mutation seam

Introduced ResolvedStep, ResolvedTarget and a closed ResolvedMutationAction. Shared49 source-free
action records and33 source-free selector records with the authored model; semantic table-row/cell
selectors remain outside the executable target seam. Only16 source-bearing mutations have distinct
resolved records, containing existing materialized core operands. Operation identity derives from
authored metadata rather than another discriminator table. ResolvedStep enforces target compatibility
and table/pivot/named-range identity; resolved grids own both collection levels and remain nonempty
and rectangular. Sheet-protection debug output redacts materialized credentials.

Production full-XSSF mutation execution now applies canonical actions directly. Streaming execution
uses the same resolved step; table-cell targets still resolve at each mutation against current
workbook state, before materialization, and secrets still resolve at the original operation stage.
Formula-origin tracking consumes resolved operands/coordinates; its real workbook tests now exercise
the new dispatcher, retaining normal/raw formula, overwrite, clear, array and append behavior.

Compile retries corrected an injected constructor call and unused match variables. Standard PMD's
NCSS rule is suppressed only on the one exhaustive65-variant dispatch, preserving one direct adapter
switch without introducing arbitrary grouping types. The initial new credential test assumed a
nonexistent defaults factory; corrected to the actual record constructor, with failure evidence kept.
Focused core/engine/parity/static checks pass (32s); new core ownership/target/redaction tests pass6s.
Final ./gradlew check passes (2m51s): core380, JSON353, engine1447 plus14 parity, CLI588, zero failures,
errors/skips; all143 golden cases unchanged. Nested check/all retained seeds passes16s. No active
mutation/fuzz campaign. Inventories: definitions0, snapshots2(runtime), converter2283, blank364,
Support201 unchanged; main Java116274→117070 during the execution cutover. The superseded command API
and its test/generator consumers remain to be deleted in the next batch; they no longer drive public
mutation execution. WP-40 is open, including ResolvedPlan and the read-side cutover. Baseline diff empty.

## WP-40 remove mutation commands and retarget deterministic replay

Deleted WorkbookCommand, its nine families, executor, discriminator table, four family converters,
selector translator and fake-grant convenience. WorkbookExecutionEngine and streaming authoring accept
ResolvedStep directly. Shared PaneSpec and SheetCopyPosition reach POI adapters without engine model
copies or reconstruction. Removed unused named-range, pane/copy and font-height converters; core
constructors/factories own the remaining contracts. Retained real mutation, persistence, selection,
protection, copy, chart and package-security integration tests; migrated their fixtures to core
coordinates/actions. Removed only retired model/translation assertions and their unused scaffolding.
Copied-sheet placement checks now live in core. Protection diagnostics remain as a separately named
test. The direct adapter integration suite is ResolvedMutationExecutorTest.

Retargeted the deterministic sequence generator, engine-sequence harness, round-trip verifier and
independent cell-footprint oracle to resolved steps. Existing seed bytes, command counts, outcomes,
style telemetry and byte-consumption order remain unchanged. Telemetry now derives the operation id
from core metadata: CREATE_SHEET becomes ENSURE_SHEET. Reviewed exactly28 promoted metadata files,
one key rename each, and two replay-test expectations; recursive comparison proves every other
field/value unchanged (external review-sequence-metadata-labels.json). Historical source paths and
failure narratives remain evidence, without runtime aliases. All94 seeds pass:48 protocol-request,
11 protocol-workflow,9 engine-sequence and26 round-trip. Nested check passes10s; no active campaigns.

Compilation/static retries removed stale imports, record accessors and private test helpers. A broad
fixture cleanup accidentally removed live diagnostic inputs; restored them from HEAD and verified the
final diff removes only the unused mutation-mirror declaration/assertions. Behavioral tests found a
real empty-batch null-workbook regression in the reduced facade; restored the boundary check. Streaming
now rejects unsupported action kinds before casting target coordinates. A copied-sheet constructor
message assertion was adapted to the canonical core constructor. Failures/logs retained externally.
Focused core/engine/parity/PMD checks pass29s. Final ./gradlew check passes4m: core382, JSON353,
engine1415 plus14 parity, CLI588, zero failures/errors/skips; all143 CLI golden cases unchanged.
Packaged engine/test-fixture/CLI/shadow JARs contain no retired command/pane/copy-position classes.

Inventories: converter lines2283→1677, blank checks364→288, Support201→200, main Java117070→114886,
test Java155456→153526; definitions0, snapshots2(runtime), core Specs75/Patches8/Reports98 unchanged.
Stopped only owned idle daemon11774 after the completed gate and zero active test executors, releasing
its4,958MiB aggregate retained worker RSS; ownership/process evidence is external. Baseline diff empty.
WP-40 remains open for ResolvedPlan, the read-side query/result/selection cutover and folding remaining
source-only operand materialization into source resolution. WP-21's under50 blank-check criterion
also remains open. No Phase4 completion or final verification claim is made here.

## WP-40 one source-bearing operand owner

Folded the six remaining WorkbookCommand-prefixed operand/source builders into MutationResolution,
then deleted those files. They materialize distinct source-bearing core contracts rather than engine
models; source-free actions still retain identity. Named the private chart title/series/plot methods
for their actual core operands. Existing source, binary, rich-text, chart, comment, table, print,
validation and secret materialization tests exercise the same owner. No replacement aliases or
compatibility classes were introduced. The preflight source-inlining stage still exists and must be
replaced by the real ResolvedPlan cutover; this commit does not claim that work is complete.

Focused engine/parity/static checks pass37s. Standard PMD initially flagged54 imports at the closed
source-bearing action boundary; documented a local suppression for this deliberately unified owner.
No global quality rule changed. Nested check/all94 retained seeds pass15s. Final ./gradlew check
passes3m; all143 CLI golden cases unchanged (142.354s), no failures/errors/skips and no active campaigns.
Inventories: converter-named lines1677→1218, Support200→199, main Java114886→114790, tests153526→153518;
blank288, definitions0 and snapshots2(runtime) unchanged. The467-line materialization owner remains
real necessary code; the inventory's converter-name reduction is not a claim that all those lines
were deleted. No WorkbookCommand reference remains in engine main code. Baseline diff empty.

## WP-40 canonical workbook/sheet summary and protection facts

Full-XSSF and event-read producers now return core WorkbookSummary, SheetSummaryReport and
SheetProtectionReport. Deleted the nested engine variants, summary validator copies and workbook-core
report converter. Read result conversion forwards these facts directly. Observed sheet labels are
constructed at native read boundaries; active/selected fallback, ordering, count and recalculation
behavior are preserved. The previous native summary's physical count/index bounds now live in the
canonical SheetSummaryReport, with core empty-sentinel/negative-bound regressions. Existing core
workbook-summary invariant tests cover the retired engine duplicate tests; removed their duplicate
state-specific constructor sweep. Real native/event/reopen/selection tests remain.

Compilation retries removed a stale static import and retired generic-string fixtures. An initial
cleanup script assumed another test method followed the last method; it failed before editing that
file and caused one repeated compile attempt. Corrected the boundary and stopped-on-error sequencing.
Eight native assertions initially compared String/List<String> with the new labels; updated fixed
expected labels without changing names or order. A subsequent overbroad assertion edit touched three
unrelated selector/grant expectations; restored those files exactly and verified them separately12s.
Failure traces are retained in the external summary-facts-* assertion JSONs. All other focused native
checks, parity and PMD pass. The independent sequence shape checks unwrap labels at their own input
boundary; all94 retained seeds and nested check pass16s. No active campaigns.

Final ./gradlew check passes3m20s: core384, JSON353, engine1414 plus14 parity, CLI588, zero failures,
errors/skips; all143 CLI golden cases unchanged. Inventories: converters1218→1160, Support199→198,
main Java114790→114625, tests153518→153591; blank288, definitions0, snapshots2(runtime) unchanged.
Baseline diff empty. WP-40 remains open for resolved plan and the remaining read/query/model cutover.
Review also found residual requireNoTables no-op wrappers and a range-only autofilter helper still
present in source, contrary to an earlier progress account; verify/remove them in the next copy cleanup.

## WP-40 remove obsolete sheet-copy reads and shims

Verified the residual requireNoTables guard has no production caller and only validates parameters;
its own test explicitly described a no-op. Removed all three forwarding layers and that test. Removed
the range-only sheetOwnedAutofilterRange helper/forwarders; retained the actual full-state selector
and its empty/sheet-owned/table-owned rejection checks. No copy protection was disabled.

The private copy plan also captured layout and merged-region facts that no copy operation consumed:
POI cloneSheet preserves those structures itself. Removed both unused fields and reads. A real
negative control copied a native hidden zero-height row and failed before the fix because the unused
layout read constructed a positive-height-only RowLayout. The same test now saves and reopens the
workbook with POI and independently confirms both source/copied row height0, hidden flag and cell text.
The change establishes copy preservation; it does not claim layout inspection now accepts zero height.
Failure XML/log retained externally. This corrects the earlier progress account that these no-op and
range-only helpers were already gone. No replacement aliases remain (source search empty).

Focused copy suite passes21s; nested check/all94 retained seeds pass14s. Final ./gradlew check passes
3m40s: core384, JSON353, engine1414 plus14 parity, CLI588, zero failures/errors/skips; all143 golden cases
unchanged. No active campaigns. Inventories: main Java114625→114592, tests153591→153596; converter1160,
blank288, Support198, definitions0, snapshots2(runtime) unchanged. Baseline diff empty. WP-40 remains
open for resolved plan and remaining read-side consolidation.

## WP-40 canonical sheet layout, merges and cell metadata facts

Deleted six nested engine fact models: SheetLayout, ColumnLayout, RowLayout, MergedRegion,
CellHyperlink and CellComment. Native layout/annotation producers now construct the existing core
reports directly. Pane reads produce PaneReport; pane writes still consume PaneSpec. The typed
address unwrap in comment repair preserves its native coordinate processing. Read result conversion
forwards these facts, deleting sheet-layout/metadata/pane reconstruction; only window conversion
remains in that helper. Native layout, saved workbook, copied-sheet and comment-column-repair checks
retain the same expected values, coordinates and order. The independent round-trip verifier reads
PaneReport from POI directly, without calling the production reader.

Compile retries corrected a comment-repair String lookup, the readback fixture's pane return type,
and an expected-map address lookup flagged by Error Prone. Initial behavioral failures were spec vs
report equality or String vs typed label/address comparisons; narrowed fixtures to factual records
and scalar values. Retired the deleted merge wrapper's blank rejection (RangeLabel intentionally
represents observed text) and the deleted row/column wrappers' wording-only checks. Existing canonical
core row/column invariants remain independently tested in WorkbookResultEdgeCoverageTest. Failure
traces retained in external sheet-facts-* JSONs. Focused fixture checks pass12s; nested check/all94
retained seeds pass16s. No active campaigns.

Final ./gradlew check passes2m58s: core384, JSON353, engine1414 plus14 parity, CLI588, zero failures,
errors/skips; all143 golden cases unchanged. Inventories: converters1160→1083, main114592→114469,
tests153596→153631; Support198, blank288, definitions0, snapshots2(runtime) unchanged. Baseline diff
empty. WP-40 remains open for ResolvedPlan, windows/surfaces/analyses and the read query/result cutover.

## WP-40 canonical analysis findings and health reports

Deleted the engine WorkbookAnalysis payload hierarchy and its health/finding/location conversion
support. Native analyzers now build existing core health reports, AnalysisFindingReport,
AnalysisSummaryReport and AnalysisLocationReport directly. Sheet labels are captured where native
locations are created; the conditional-formatting sheet-name lookup explicitly unwraps its label.
Read result conversion forwards analysis payloads unchanged. Retired only the deleted engine DTO's
two constructor/collection-copy tests; existing canonical core invariant tests and real analyzer,
selection, evidence, severity, count and ordering tests remain.

Main compilation identified one String boundary needing label.value(); corrected it. Focused core,
engine, parity and PMD checks found only a String-vs-label assertion in the selected-sheet validation
health test; retained its exact expected sheet and counts, then verified that class16s. Failure XML
retained externally. Nested check/all94 retained seeds passes19s. Final ./gradlew check passes3m4s:
core384, JSON353, engine1412 plus14 parity, CLI588, zero failures/errors/skips; all143 golden cases
unchanged. No active campaigns. Inventories: converters1083→927, Support198→197, blank288→287,
main Java114469→114102, tests153631→153526; definitions0 and snapshots2(runtime) unchanged. Baseline
diff empty. WP-40 remains open for ResolvedPlan, surface/window facts and the read query/result seam.

## WP-40 canonical formula, schema and named-range surfaces

Deleted eight nested engine surface records and the duplicate NamedRangeBackingKind enum. Native
formula grouping, schema inference and named-range classification now construct the existing core
reports and enum directly; typed sheet labels and formula-cell addresses are created at read
boundaries. Deleted the surface report converter. Removed unused projection/date-system fields from
SheetSchemaResult; the producer still consumes the requested projection and resolves cell temporal
facts before inference. No schema counts, grouping/classification or selection algorithm changed.

Main compilation passes14s. Test compilation found one qualified type split across lines; corrected
that stale reference. Focused behavior found two String/label assertions and the deleted native
schema-column wrapper's wording checks. Preserved the exact expected values and retired those
wording-only assertions; canonical schema invariants remain in WorkbookResultEdgeCoverageTest.
Failure traces retained in surface-facts-fixture-failures.json. Final focused fixture checks, PMD,
parity and nested check/all94 seeds pass. No active campaigns. Final ./gradlew check passes3m38s:
core384, JSON353, engine1412 plus14 parity, CLI588, zero failures/errors/skips; all143 golden cases
unchanged. Inventories: converters927→829, Support197→196, main114102→113870, tests153526→153527;
blank287, definitions0, snapshots2(runtime) unchanged. Baseline diff empty. WP-40 remains open for
ResolvedPlan, physical windows and the remaining read query/result/selection seam.

## WP-40 return core inspection results directly

Deleted the complete engine read-result hierarchy, the remaining physical Window/WindowRow mirrors,
its validation helper and all result-conversion classes. Full-XSSF, event-read, assertions and runtime
inspection flows now receive canonical InspectionResult variants directly. Physical cell reads build
WindowReport.Dense/WindowRowReport; the query boundary chooses dense or sparse output and omits blank
reports for sparse queries with the same prior semantics. Projection and date-system calculation stay
at native cell read boundaries. No result-conversion alias remains. The independent sequence replay
and round-trip consumers now use the same canonical result types.

Removed only the three fixture-only result/model conversion test classes and two corresponding methods
from mixed tests, plus their unused private fixtures. Existing core result invariants and real native,
public workflow, event-read, assertion, host-acceptance, persistence and golden tests remain. Added a
real query regression: sparse returns exactly A1/B2/B3, dense retains the3×2 grid and B1's blank-cell
comment, and both retain their step ids. The defaults were preserved: an existing physical read test
now explicitly observes its requested sparse form instead of expecting a raw intermediate grid.

Main compilation passes8s. Compile retries corrected print-layout/window accessors, a qualified-name
rewrite and unused private test helpers. PMD's import-count rule crossed its threshold by moving the
closed analysis dispatcher's result type into core; documented a local suppression, without a global
rule change or artificial facade. Focused core/engine/parity/PMD checks pass32s and nested check/all94
retained seeds pass17s. Final ./gradlew check passes2m55s: core384, JSON353, engine1394 plus14 parity,
CLI588, zero failures/errors/skips; all143 golden cases unchanged. No active campaigns. Four packaged
JARs were independently checked: all14 deleted main-type families are absent, including nested classes.

Inventories: converters829→625, Support196→194, blank287→286, main113870→113122, tests153527→150937;
definitions0, snapshots2(runtime) unchanged. Only the historical changelog archive mentions the retired
read converter; current Markdown guidance has no such reference. Baseline diff empty. WP-40 remains
open for ResolvedPlan and elimination of the query/selection translation seam (the last625 converter
lines), before feature-package/backend work.
