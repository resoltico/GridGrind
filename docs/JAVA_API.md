# Java API

Pure authoring lives in core.authoring. GridGrindPlan and the selector-first Targets, Values,
ExpectedValues, Tables and Links facades create the same canonical WorkbookPlan used by JSON.
They do not create a second execution engine or implicit authority.

The [compile-verified example](../examples/java-authoring-workflow.java) demonstrates typed
SheetName, RangeRef and TableName values, source-backed text, ordered operations and explicit
execution bindings. [Model vocabulary](../core/VOCABULARY.md) distinguishes authored specs,
patches, factual reports and source-bearing resolved values.

Execution belongs to engine.api. Obtain the request executor or doctor from GridGrindEngine
and supply host-owned GridGrindRequestInputs with an ExecutionGrant. Applications that only
author or serialize plans need core/json, without the workbook runtime. Execution adds engine.

Optional fields use Optional and are omitted from JSON when absent. Wire scalar value records
serialize as JSON primitives, preserving identifier strings. Typed values validate their
authoring or observation contract; observed labels do not silently acquire authoring restrictions.

Use [generated reference](reference/README.md), [execution and grants](EXECUTION_MODEL.md),
and [development](DEVELOPER.md) for the exact model and verification commands.
