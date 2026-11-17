# Model vocabulary

`Spec` names a complete authored choice or definition. A spec can have declared omission defaults;
its effective value is complete after those defaults are applied. Text and binary sources are specs,
and their standard-input variant is `Stdin`.

`Patch` names a partial authored update. Its fields are optional and the constructor rejects an empty
update. Cell style, font, alignment, border, border-side and protection patches use this convention.
A fill spec describes one chosen fill operation inside a style patch; its selected variant has the
required operands for that operation.

`Report` names observed facts. A field can retain a stored declaration or expose an effective
value according to its contract. Authored identifiers and observed labels have distinct
contracts: for example, `SheetName` enforces worksheet-authoring rules while `SheetLabel` preserves
an observed name, and `RangeLabel` preserves blank stored text and unsupported references such as `#REF!`.
`CellAddress` stores zero-based row and column indices; `RangeRef` stores an ordered rectangle.
Their A1 presentation is canonical: case and `$` markers do not affect concrete target identity,
and `A1:A1` has the same identity as `A1`. Rectangles reject reversed endpoints. Selector lists reject
equivalent coordinates and retain ordered, nonidentical overlaps. Formula text and observed labels
retain their own spelling and anchoring contracts.
Named-range scope specs use authored `SheetName`; `NamedRangeScopeReport` retains observed
`SheetLabel` values. Formula-backed name reports can preserve an empty stored reference.

Model types do not use `Input`, `Dto`, `Snapshot` or `Definition` as generic suffixes. Materialized
source-bearing specifications use the `Resolved` prefix where a separate resolved form is necessary.

Prepared steps retain original authored steps separately from bound selectors and operands.
Preparation permits only declared external-source binding; unrelated actions, queries, selectors
and unresolved external sources reject. A prepared value does not grant execution authority;
engine preflight checks the host grant before issuing its private execution capability.

A union's identity is declared by `Union` and `Variant` metadata. Its catalog group id is the union's
simple name with a lower-case initial plus `Types`. A plain catalog type uses its declared id with a
lower-case initial plus `Type`. The catalog and codec derive identities from that source.

Collections use plural domain nouns. Optional values use `Optional<T>`; absent values are omitted
from the wire.
`WireDefault` declares omission defaults; an explicit JSON `null` is not an omission default.

Wire discriminators use `type` for a variant's identity and `status` for a completed process outcome.
The structural `WorkbookStep` envelope selects its action, assertion or query without an additional
tag. Variant ids use `SCREAMING_SNAKE_CASE` and are declared by `Variant` metadata. Plain attributes
such as journal verbosity and a diagnostic cause's stage keep their domain names.

`protocolVersion` is the integer `1`. Validated wire scalars represent domain values directly as
JSON primitives; their Java record components do not add an object layer to the wire document.

`CellValueSpec` describes authored cell content. Each scalar or source-bearing variant uses
`value`; rich text uses `runs`, whose source-bearing text field is `text`. `CellValue` is the
shared scalar union for observations and exact expectations: blank has no payload,
and text, number, boolean and error variants each use `value` for their primitive payload.
Observed text may be empty and observed numbers may be nonfinite; authored numeric expectations
remain finite. Exact text and evaluation-only error expectations retain their observation semantics.
`ResolvedCellWrite` accepts validated write operands: finite authored numbers, stored errors,
materialized XML-safe cell text, formulas, rich text and temporal values. Blank and boolean leaves
are shared where their guarantees match. Other observed scalars cannot be passed to a writer.
`CellReport` retains the stored cell type and carries the shared scalar in its optional `value`
field. Formula text, rich-text runs and temporal metadata stay on the report and follow its
projection facets. A formula report's scalar value describes the evaluated type.

`SecretReference` retains its object shape with a required `id`; it is a source reference rather
than a `WireScalar`. Its catalog and schemas describe that same object and its constructor-owned
identifier grammar. Secret material stays outside the document model.
