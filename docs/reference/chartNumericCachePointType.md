<!-- Generated from core metadata; do not edit. -->
# chartNumericCachePointType

## ChartNumericCachePoint

One stored numeric-cache point with raw value text and optional per-point format. Missing, blank, error and nonfinite-looking text remain distinct facts.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| index | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":0.0,"maximum":4.294967295E9}]} |  |
| valueText | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| formatCode | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "ChartNumericCachePoint",
  "summary" : "One stored numeric-cache point with raw value text and optional per-point format. Missing, blank, error and nonfinite-looking text remain distinct facts.",
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
  }, {
    "name" : "formatCode",
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
