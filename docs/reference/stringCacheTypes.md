<!-- Generated from core metadata; do not edit. -->
# stringCacheTypes

## MISSING

No string cache element is stored.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "MISSING",
  "summary" : "No string cache element is stored.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## STORED

Stored string cache with its optional declared count and physical point order.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| declaredCount | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":0.0,"maximum":4.294967295E9}]} |  |
| points | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartStringCachePointType"}} |  |

```json
{
  "id" : "STORED",
  "summary" : "Stored string cache with its optional declared count and physical point order.",
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
    "name" : "points",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartStringCachePointType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
