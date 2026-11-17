<!-- Generated from core metadata; do not edit. -->
# chartAxisReportType

## ChartAxisReport

Exact chart-axis report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| kind | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| position | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| crosses | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"axisCrossingTypes"} |  |
| visible | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "ChartAxisReport",
  "summary" : "Exact chart-axis report.",
  "fields" : [ {
    "name" : "kind",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "CATEGORY", "DATE", "SERIES", "VALUE" ]
  }, {
    "name" : "position",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "BOTTOM", "LEFT", "RIGHT", "TOP" ]
  }, {
    "name" : "crosses",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "axisCrossingTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "visible",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
