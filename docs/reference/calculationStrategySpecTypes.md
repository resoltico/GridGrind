<!-- Generated from core metadata; do not edit. -->
# calculationStrategySpecTypes

## DO_NOT_CALCULATE

Skip immediate server-side formula calculation.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "DO_NOT_CALCULATE",
  "summary" : "Skip immediate server-side formula calculation.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## DEFERRED_CALCULATION

Inspect formula capabilities and defer all calculation to an Excel-compatible client.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "DEFERRED_CALCULATION",
  "summary" : "Inspect formula capabilities and defer all calculation to an Excel-compatible client.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## EVALUATE_ALL

Preflight and evaluate every formula cell after mutation steps complete.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "EVALUATE_ALL",
  "summary" : "Preflight and evaluate every formula cell after mutation steps complete.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## EVALUATE_TARGETS

Preflight and evaluate the explicit formula-cell target list only.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"formulaCellTargetType"}} |  |

```json
{
  "id" : "EVALUATE_TARGETS",
  "summary" : "Preflight and evaluate the explicit formula-cell target list only.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "formulaCellTargetType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## REQUIRE_EVALUATION

Preflight and evaluate every formula cell, failing when any formula cannot be evaluated immediately.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "REQUIRE_EVALUATION",
  "summary" : "Preflight and evaluate every formula cell, failing when any formula cannot be evaluated immediately.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## CLEAR_CACHES_ONLY

Strip persisted formula caches without running immediate evaluation.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "CLEAR_CACHES_ONLY",
  "summary" : "Strip persisted formula caches without running immediate evaluation.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```
