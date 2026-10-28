# Execution model

A WorkbookPlan describes intended work. A host-owned ExecutionGrant independently permits
read resources, operation IDs, workbook target scope, publication and secret references.
Plan content cannot manufacture authority. Java and CLI execution use the same grant model.

Admission validates structure and defaults before binding request-owned paths. The JSON
pipeline decodes UTF-8, parses once with recovery, collects schema/domain findings, materializes
omission defaults, then binds the validated model. A binding failure after validation is an
internal defect. Doctor retains safe independent fragments and reports every provable problem
without applying workbook mutations.

## Paths and secrets

File requests establish their containing directory as the execution root; stdin requests need
an explicit root. Relative and absolute request-owned paths must remain under that root.
Descriptor-relative no-follow access rejects symlinks and unsafe ancestors. Grants authorize
exact resources rather than bypassing those protections. Response output is host transport.

Secret references are resolved by the trusted host and separately allowed by the grant.
Diagnostic redaction covers authored secret-capable input. Encrypted package plaintext and
run scratch have explicit owners and cleanup lifetimes.

## Execution

Specs carry complete authored choices, patches carry partial updates, reports carry observed
facts, and resolved types exist where source binding changes the contract. Selectors resolve
to concrete targets once. FULL_XSSF performs the complete in-memory workflow. EVENT_READ and
STREAMING_WRITE have explicit reduced operation capabilities; unsupported operations are rejected.
EVENT_READ uses its direct streaming reader for eligible existing-source requests without publication.
Existing-source preflight opens full XSSF first, and other admitted EVENT_READ requests use XSSF
materialization. STREAMING_WRITE uses SXSSF for append authoring, but assertions, inspections
and staged-publication verification can reopen full XSSF. Neither mode guarantees bounded
end-to-end heap use.

Steps preserve authored order. Assertions may collect after the terminal assertion phase
begins; no later mutation is legal. Inspections may interleave. Factual observations retain
stored labels and unsupported details without authoring them as new valid identifiers.

## Publication and evidence

Publication requires matching host authority. The engine stages an artifact, checks acceptance,
then publishes through the admitted descriptor/collision policy. The response distinguishes
NOT_ATTEMPTED, proven NOT_PUBLISHED, PUBLISHED and UNCERTAIN destination outcomes. An UNCERTAIN
outcome requires inspecting the destination before retrying. A failed execution alone does not
prove that the destination is unchanged. Existing-source security dispositions are explicit.
Response publication is separate from workbook publication; if it fails, the already-rendered
primary payload can fall back to stdout with a structured stderr notice.

Journals record execution evidence. VERBOSE progress goes to the diagnostic stream rather than
duplicating events in the primary response. Runtime before/after acceptance snapshots, package
relation snapshots and unsupported native-style acquisition snapshots are operational state,
not duplicate core model types.

See [limits](LIMITATIONS.md), [problems](ERRORS.md), [CLI](CLI.md) and the
[generated reference](reference/README.md) for exact fields and supported variants.
