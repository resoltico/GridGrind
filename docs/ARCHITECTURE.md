# Architecture

The product modules are core, json, engine and cli. Architecture-tests verifies their
compiled boundaries; fuzz is a root test source set.

    core <- json <- cli
      ^---- engine <- cli
    architecture-tests -> product modules

Core is POI-free and Jackson-free. It owns vocabulary, validated scalar values, authored
specs, patches, factual reports, source-bearing resolved values, executable plans, grants,
semantic rules, discovery and metadata. JSON derives union registration and schemas from
that metadata, while its tolerant parser retains syntax findings and independently valid
fragments for complete diagnosis.

Engine admits a plan, binds only authorized inputs, resolves semantic selectors into concrete
targets and executes core operations. Native POI/OOXML mechanics stay in feature and backend
packages. Only engine.api is exported. Execution backends consume the canonical plan and
publish capabilities from one pure declaration; routing, validation and help use that declaration.

The CLI contains command, render and IO packages. Pure recipes and matching belong to
core.discovery; workspace materialization is a host responsibility. Source-backed inputs
become resolved values before execution. The engine has no separate public model mirror.

JPMS enforces readability and exports. ArchUnit protects native-type confinement, generic API
signatures and package/lifecycle boundaries, with independent negative controls. Runtime tests,
POI parity, property tests, the golden corpus and deterministic fuzz replay exercise behavior.
Coverage and mutation scores are advisory.

[Model vocabulary](../core/VOCABULARY.md), [execution](EXECUTION_MODEL.md),
[development](DEVELOPER.md) and [generated reference](reference/README.md) describe the
supported contracts and verification entrypoints.
