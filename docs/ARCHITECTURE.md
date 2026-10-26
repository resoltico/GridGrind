# Architecture

Core owns the canonical request/result vocabulary, metadata, static validation and Java authoring
API. It has no POI or engine dependency. JSON owns tolerant intake and wire codecs. Engine and CLI
consume the same core contracts. Execution requires an explicit host grant.

The current Gradle/JPMS graph is:

```text
core <- engine <- cli
  ^------ json <---|
architecture-tests -> product projects (verification only)
```

`engine` exports only `dev.erst.gridgrind.engine.api`. Its runtime admits plans, binds approved
inputs, resolves selectors, executes workbook operations, collects evidence and publishes verified
staged artifacts. The CLI handles arguments, transport, rendering and workspace materialization.

Production module descriptors enforce readability and encapsulation. Architecture tests import
compiled product roots from Gradle and protect API signatures, intra-engine dependencies,
closed wire variants, centralized POI access and execution/publication boundaries. Negative
controls demonstrate those checks detect violations. Runtime and XSSF parity tests live in engine.

The golden corpus executes every shipped example and recipe through in-process and packaged
CLI transports, then reopens persisted workbook facts through product inspection. Coverage and
mutation reports are advisory; behaviour tests, properties and seed replay establish assurance.

JSON emits Draft 2020-12 schemas for requests, execution grants and responses from canonical
record components, union metadata, field constraints and omission defaults. Run
`./gradlew :json:generateSchemas` to update [the reference schemas](reference/schema/).
`JsonSchemas.document` exposes the same immutable documents to Java callers. `default` carries
literal omission values; `x-gridgrind-defaultFrom` identifies a sibling-derived default. Standard
schema validation covers structure and declared scalar constraints. Canonical constructors, numeric
token checks and static/runtime validation additionally enforce domain and workbook-state rules.
Independent NetworkNT tests validate examples and actual serialized grants/results, reject malformed
controls, and compare the committed schema artifacts without rewriting them. Their reference and
example directories are declared test inputs.
