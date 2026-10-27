<!-- Generated from core metadata; do not edit. -->
# executionModeSpecTypes

## FULL_XSSF

Standard full workbook read and write execution with no special low-memory restrictions.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "FULL_XSSF",
  "summary" : "Standard full workbook read and write execution with no special low-memory restrictions.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## EVENT_READ

GET_WORKBOOK_SUMMARY and GET_SHEET_SUMMARY only, with execution.calculation.strategy=DO_NOT_CALCULATE and markRecalculateOnOpen=false.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "EVENT_READ",
  "summary" : "GET_WORKBOOK_SUMMARY and GET_SHEET_SUMMARY only, with execution.calculation.strategy=DO_NOT_CALCULATE and markRecalculateOnOpen=false.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## STREAMING_WRITE

source.type must be NEW; mutation actions limited to ENSURE_SHEET and APPEND_ROW; execution.calculation must keep strategy=DO_NOT_CALCULATE and may set markRecalculateOnOpen=true.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "STREAMING_WRITE",
  "summary" : "source.type must be NEW; mutation actions limited to ENSURE_SHEET and APPEND_ROW; execution.calculation must keep strategy=DO_NOT_CALCULATE and may set markRecalculateOnOpen=true.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```
