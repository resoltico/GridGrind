# Model vocabulary

`Spec` names a complete authored choice or definition. A spec can have declared omission defaults;
its effective value is complete after those defaults are applied. Text and binary sources are specs,
and their standard-input variant is `Stdin`.

`Patch` names a partial authored update. Its fields are optional and the constructor rejects an empty
update. Cell style, font, alignment, border, border-side and protection patches use this convention.
A fill spec describes one chosen fill operation inside a style patch; its selected variant has the
required operands for that operation.

`Report` names observed effective state. Authored identifiers and observed labels have distinct
contracts: for example, `SheetName` enforces worksheet-authoring rules while `SheetLabel` preserves
an observed name, and `RangeLabel` preserves blank stored text and unsupported references such as `#REF!`.
Authored `RangeRef` still rejects blank and malformed coordinates.

Model types do not use `Input`, `Dto`, `Snapshot` or `Definition` as generic suffixes. Materialized
source-bearing specifications use the `Resolved` prefix where a separate resolved form is necessary.

A union's identity is declared by `Union` and `Variant` metadata. Its catalog group id is the union's
simple name with a lower-case initial plus `Types`. A plain catalog type uses its declared id with a
lower-case initial plus `Type`. The catalog and codec derive identities from that source.

Collections use plural domain nouns. Optional fields use `Optional<T>` and are absent on the wire.
`WireDefault` declares omission defaults; an explicit JSON `null` is not an omission default.

Wire discriminators use `type` for a variant's identity and `status` for a completed process outcome.
The structural `WorkbookStep` envelope selects its action, assertion or query without an additional
tag. Variant ids use `SCREAMING_SNAKE_CASE` and are declared by `Variant` metadata. Plain attributes
such as journal verbosity and a diagnostic cause's stage keep their domain names.

`protocolVersion` is the integer `1`. Validated wire scalars represent domain values directly as
JSON primitives; their Java record components do not add an object layer to the wire document.

`CellValueSpec` describes authored cell content. Each scalar or source-bearing variant uses
`value`; rich text uses `runs`, whose source-bearing text field is `text`. `CellValue` is the
shared resolved scalar union for observations and exact expectations: blank has no payload,
and text, number, boolean and error variants each use `value` for their primitive payload.
`CellReport` retains the stored cell type and carries the shared scalar in its optional `value`
field. Formula text, rich-text runs and temporal metadata stay on the report and follow its
projection facets. A formula report's scalar value describes the evaluated type.

`SecretReference` retains its object shape with a required `id`; it is a source reference rather
than a `WireScalar`. Its catalog and schemas describe that same object and its constructor-owned
identifier grammar. Secret material stays outside the document model.
