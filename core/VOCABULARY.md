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
an observed name, and `RangeLabel` can retain an unsupported stored reference such as `#REF!`.

Model types do not use `Input`, `Dto`, `Snapshot` or `Definition` as generic suffixes. Materialized
source-bearing specifications use the `Resolved` prefix where a separate resolved form is necessary.

A union's identity is declared by `Union` and `Variant` metadata. Its catalog group id is the union's
simple name with a lower-case initial plus `Types`. A plain catalog type uses its declared id with a
lower-case initial plus `Type`. The catalog and codec derive identities from that source.

Collections use plural domain nouns. Optional fields use `Optional<T>` and are absent on the wire.
`WireDefault` declares omission defaults; an explicit JSON `null` is not an omission default.
