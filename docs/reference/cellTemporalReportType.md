<!-- Generated from core metadata; do not edit. -->
# cellTemporalReportType

## CellTemporalReport

Projected temporal facet for one numeric cell value, including whether Excel formatting marks it as date-like and any derived DATE, TIME, or DATE_TIME semantic kind.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| isDate | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| kind | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| isoValue | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "CellTemporalReport",
  "summary" : "Projected temporal facet for one numeric cell value, including whether Excel formatting marks it as date-like and any derived DATE, TIME, or DATE_TIME semantic kind.",
  "fields" : [ {
    "name" : "isDate",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "kind",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "DATE", "TIME", "DATE_TIME" ]
  }, {
    "name" : "isoValue",
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
