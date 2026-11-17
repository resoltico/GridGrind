<!-- Generated from core metadata; do not edit. -->
# numericCacheTypes

## MISSING

No numeric cache element is stored.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "MISSING",
  "summary" : "No numeric cache element is stored.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## STORED

Stored numeric cache. Count, formats and raw point text are never inferred from live cells.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| declaredCount | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":0.0,"maximum":4.294967295E9}]} |  |
| formatCode | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| points | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartNumericCachePointType"}} |  |

```json
{
  "id" : "STORED",
  "summary" : "Stored numeric cache. Count, formats and raw point text are never inferred from live cells.",
  "fields" : [ {
    "name" : "declaredCount",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "NUMBER_RANGE",
        "minimum" : 0.0,
        "maximum" : 4.294967295E9
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "formatCode",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "points",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartNumericCachePointType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
