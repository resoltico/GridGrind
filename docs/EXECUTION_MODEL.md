# Execution model

The engine executes admitted, source-materialized plans through one of three backends.
Native workbook mechanics live in feature adapters and the XSSF, streaming and event-read
packages. Runtime owns admission, journal, assertions, host acceptance and publication.

## Routing and constraints recorded before the backend cutover

The full-XSSF workflow operates on an opened workbook, executes mutations and observations,
performs the calculation and host-acceptance phases, persists, and closes the workbook.
It supports the complete operation surface. It also supplies the materialized path for an
EVENT_READ request that cannot take the direct source-read path.

Direct event-read is selected only for EVENT_READ with DO_NOT_CALCULATE, no recalculate-on-open,
inspection-only steps, an existing-file source and no persistence. Static validation permits
GET_WORKBOOK_SUMMARY and GET_SHEET_SUMMARY only and rejects assertions and mutations. When direct
routing is unavailable, the XSSF workflow serializes a private plain workbook for the event
reader; the event-reader query restriction still applies. Encrypted sources are decrypted into
an owned private workbook before reading, and their cleanup ownership survives through closure.

Streaming-write requires a NEW source, DO_NOT_CALCULATE, and mutations limited to ENSURE_SHEET
and APPEND_ROW. Recalculate-on-open is allowed. An ENSURE_SHEET must precede append, inspection or
assertion, and a streaming plan must create at least one sheet. Observations and assertions use
materialized private workbooks; the streaming writer does not provide arbitrary in-place edits.
Host acceptance verifies the materialized terminal workbook before publication.

These observations come from the existing `ExecutionWorkflowRouting`, `FullXssfWorkflow`,
`ExecutionStreamingWorkflow`, `ExecutionDirectEventReadWorkflow`, `StepExecution`,
`WorkbookStaticExecutionValidation`, `WorkbookStaticStreamingValidation`,
`WorkbookOperationContract`, and `CalculationPolicySpec` implementations.

## Backend capability ownership

Capability declarations are pure core semantics, shared by the execution backends, static
validation and catalog/help rendering. Each backend exposes its canonical declaration through
`ExecutionBackend`; runtime routing reads it. The declarations include calculation restrictions
and the direct event-read route, so consumers do not maintain separate policy lists.

The capability contract must live in core: core owns validation and catalog and cannot depend
on engine. This is the layering correction to WP-42's literal engine-owned declaration wording;
there is one shared declaration per mode, with no reverse module dependency or copied registry.
