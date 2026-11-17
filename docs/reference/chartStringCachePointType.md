<!-- Generated from core metadata; do not edit. -->
# chartStringCachePointType

## ChartStringCachePoint

One stored string-cache point. Missing index/value declarations stay absent; list position does not infer the index.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| index | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":0.0,"maximum":4.294967295E9}]} |  |
| valueText | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "ChartStringCachePoint",
  "summary" : "One stored string-cache point. Missing index/value declarations stay absent; list position does not infer the index.",
  "fields" : [ {
    "name" : "index",
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
    "name" : "valueText",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
