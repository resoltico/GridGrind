# Execution model

A WorkbookPlan describes intended work. A host-owned ExecutionGrant independently permits
read resources, a finite aggregate input-snapshot byte budget, operation IDs, workbook target
scope, publication and secret references.
Plan content cannot manufacture authority. Java and CLI execution use the same grant model.

Admission validates structure and defaults before binding request-owned paths. The JSON
pipeline decodes UTF-8, recovers syntax and independent structural findings, and constructs each
valid canonical fragment once. Omission defaults apply to legitimately absent fields; malformed,
duplicate-owned and explicit-null fields cannot acquire executable values through recovery.
Doctor retains safe independent fragments without applying workbook mutations. Static plan refusals
stop acquisition; after static admission, authorized input and source checks retain independent
findings. The sole preparation factory enforces static rules before opening inputs, independently
of the caller's validation-phase telemetry.

## Paths and secrets

File requests establish their containing directory as the execution root; stdin requests need
an explicit root. Relative and absolute request-owned paths must remain under that root.
Descriptor-relative no-follow access rejects symlinks and unsafe ancestors. Grants authorize
exact resources rather than bypassing those protections. Filesystem-native case and Unicode
aliases use verified stored directory-entry names consistently for containment, grant matching and
snapshot reuse. Separate hard-link names require separate grants and consume separate snapshot
bytes. An ambiguous non-exact alias refuses; use its exact stored name. Missing output suffixes
retain their exact authored spelling. Unused nonexistent or outside-root grants grant no access and
do not invalidate other grants. Trusted ancestors above the host root need only traversal for
ordinary contained paths; proving an alternate absolute root spelling also requires readable
directory metadata. Response output is host transport.

`maxInputSnapshotBytes` is required, finite and nonnegative. It covers source workbooks,
formula workbooks, FILE/UTF8_FILE payloads and standard input, including source-backed host
acceptance selectors. Actual bytes are charged while snapshotting; aliases of one admitted
resource and repeated stdin uses share one immutable snapshot. Zero permits NEW/inline work
when the grant authorizes its operations. An overflow or failed snapshot read blocks further reads in that run; already
completed snapshots remain reusable. Overflow consumes the remaining allowance, with at most one
lookahead byte read to detect it. The allowance does not bound decompression, native
workbook memory, output growth or general scratch capacity. Authority documents reject duplicate,
unknown, malformed and explicit-null fields and trailing roots as a whole. Invalid authority is
`INVALID_AUTHORITY` with `READ_AUTHORITY` context, the selected grant file and reliable JSON
coordinates. CLI run and doctor reject it before workbook execution with exit 2; they accept no
partial or default authority. A valid insufficient grant instead yields `AUTHORITY_DENIED`.
Authority file access failures remain I/O failures; unknown reader/provider defects remain internal.

Secret references are resolved by the trusted host and separately allowed by the grant.
The CLI file provider reads each no-follow leaf through EOF, rejecting more than 65,536 bytes
and malformed UTF-8. File size is only an early hint; growth cannot bypass the actual read limit.
Diagnostic redaction covers authored secret-capable input. Encrypted package plaintext and
run scratch have explicit owners and cleanup lifetimes. Native execution uses a private provider realm for each invocation, with fixed en-US rendering
and a UTC civil-date bridge. Embedding host locale, time zone and provider settings remain owned
by the host. Run-owned native scratch stays within its admitted private root. Cleanup failures retain an existing primary problem and completed execution evidence.

## Authored media

Picture, embedded-object preview and signature-line preview inputs are checked before native
mutation. Declared formats select the engine's bounded Java structural reader; no external
image utility or PostScript interpreter is required. Format signatures alone do not establish
admission. Nested raster data, coding, declared spans and format-specific relationships must
satisfy the supported profile.

The [media limits](LIMITATIONS.md#lim-042--authored-media-payload-bytes) separately bound one
payload, authored uses, inspection, raster pixels, expansion, structural records, nesting and
EPS parsing work. Repeated authored uses charge even when source bytes share one input
snapshot. These controls do not certify a renderer's visual output or bound all workbook heap
use. Existing observed media and newly authored operands retain separate contracts.

## Execution

Specs carry complete authored choices, patches carry partial updates, reports carry observed
facts, and resolved types exist where source binding changes the contract. Selectors resolve
to concrete targets once. Prepared steps retain original authored sources and occurrence locations
separately from bound operands; failure evidence reports original selectors alongside observations.
DATE and DATE_TIME are Gregorian civil values, without an inferred time zone. Authoring uses
1900-01-01 through 9999-12-31 for a 1900 workbook, or 1904-01-01 through 9999-12-31 for a
1904 workbook. DATE_TIME accepts millisecond precision; finer fractions refuse rather than
rounding silently. Intrinsic temporal operands are checked after the actual source calendar is
known and before plan mutations. A mapped XML DATE is checked against the bindings used by
that import, before its first cell, row or table mutation; missing mappings retain their no-op
semantics. Mapped dates must be valid, timezone-free XML Schema dates. An explicit offset is
refused, never discarded.

TEMPORAL is an optional interpretation beside the factual numeric VALUE and format. Its closed
outcomes are NOT_TEMPORAL, INTERPRETED and UNREPRESENTABLE. Unsupported serials, the fictitious
1900 leap day and elapsed formats retain the original number without acquiring a fabricated ISO
date. Time-only interpretation uses the nonnegative fractional day; elapsed hours are not wrapped
into a clock time. Request VALUE alongside TEMPORAL when the original number is needed.

Calculation, display assertions, normal and cached-formula formatting, and deterministic column
sizing use the same fixed en-US policy. A spreadsheet civil time is not converted through a host
DST gap. UTC serves only as the neutral bridge for legacy date APIs. No request locale/timezone
field or fallback policy is implied.

FULL_XSSF performs the complete in-memory workflow. EVENT_READ and
STREAMING_WRITE have explicit reduced operation capabilities; unsupported operations are rejected.
Admission selects the concrete backend and its required full-workbook views once. Eligible
EVENT_READ requests against an existing source without publication validate package security,
passwords and workbook metadata through the event path, retaining the readable source for
execution. FULL_XSSF retains its prepared native source and formula environment for execution;
other admitted EVENT_READ requests require XSSF materialization.

STREAMING_WRITE uses SXSSF for append authoring followed by terminal observations, which share
one serialized snapshot and full XSSF view. A plan that mutates after an observation selects the
reported FULL_XSSF backend so each observation sees its authored state; that plan requires full
workbook memory. Streaming source, operation and calculation restrictions still apply. Host terminal assertions share one final artifact view, and
persisted output is reopened through full XSSF for staged-artifact verification. Required host
acceptance is never dropped to preserve a mode label. Neither mode guarantees bounded
end-to-end heap use.

Result admission evidence reports the selected `executionStrategy.backend` and whether
`fullWorkbookMaterializationRequired` is true. This describes the admitted requirements; it does
not claim that every required stage completed successfully. An absent strategy means no strategy
was admitted. Structural, calculation, assertion and publication evidence report established facts
separately.

Steps preserve authored order. Assertions may collect after the terminal assertion phase
begins; no later mutation is legal. Inspections may interleave. Factual observations retain
stored labels and unsupported details without authoring them as new valid identifiers.

## Publication and evidence

Publication requires matching host authority. The engine stages and verifies a complete artifact
in a private destination sibling. On qualified JDK filesystem providers, create-only publication
uses atomic hard-link creation; replacement uses atomic rename. Neither path exposes an incomplete
final artifact. The destination parent and ancestors must remain trusted and stable through the
pathname-based commit; descriptor-relative reads and topology rechecks do not close the final
pathname scheduling gap. Unsupported required primitives fail closed.

NOT_ATTEMPTED means the final mutation primitive was never invoked. NOT_PUBLISHED proves that
this run did not commit; COLLISION_OBSERVED records occupancy at a qualified create-only primitive,
not the destination's later state. PUBLISHED requires verified complete content and file-sync
evidence, with directory-sync evidence reported separately. UNCERTAIN preserves whether the
primitive returned successfully and requires inspection before retrying. Identical destination
bytes alone never establish this run's commit. Cleanup and optional progress failures cannot erase
known publication occurrence or completed assertion and inspection results.
Response publication is separate from workbook publication; if it fails, the already-rendered
primary payload can fall back to stdout with a structured stderr notice.

Journals record execution evidence. VERBOSE progress goes to the diagnostic stream rather than
duplicating events in the primary response. A progress-sink RuntimeException disables that sink
once and adds one structured warning; caller Error propagates after resource unwinding. One final
plan outcome is emitted after required session cleanup. Runtime before/after acceptance snapshots, package
relation snapshots and unsupported native-style acquisition snapshots are operational state,
not duplicate core model types.

See [limits](LIMITATIONS.md), [problems](ERRORS.md), [CLI](CLI.md) and the
[generated reference](reference/README.md) for exact fields and supported variants.
