<!-- Generated from core metadata; do not edit. -->
# phaseTypes

## NOT_STARTED

Execution phase that never began and therefore exposes no timing payload.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NOT_STARTED",
  "summary" : "Execution phase that never began and therefore exposes no timing payload.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## NOT_REQUESTED

Execution phase intentionally skipped and therefore exposing no timing payload.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NOT_REQUESTED",
  "summary" : "Execution phase intentionally skipped and therefore exposing no timing payload.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## SUCCEEDED

Execution phase that completed successfully and may include one timing payload.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| timing | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"executionJournalTimingType"} |  |

```json
{
  "id" : "SUCCEEDED",
  "summary" : "Execution phase that completed successfully and may include one timing payload.",
  "fields" : [ {
    "name" : "timing",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "executionJournalTimingType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## FAILED

Execution phase that ended in failure and may include one timing payload.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| timing | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"executionJournalTimingType"} |  |

```json
{
  "id" : "FAILED",
  "summary" : "Execution phase that ended in failure and may include one timing payload.",
  "fields" : [ {
    "name" : "timing",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "executionJournalTimingType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
