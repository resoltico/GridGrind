<!-- Generated from core metadata; do not edit. -->
# cellGradientStopReportType

## CellGradientStopReport

Exact gradient stop report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| position | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| color | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellColorReportTypes"} |  |

```json
{
  "id" : "CellGradientStopReport",
  "summary" : "Exact gradient stop report.",
  "fields" : [ {
    "name" : "position",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "color",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellColorReportTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
