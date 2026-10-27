<!-- Generated from core metadata; do not edit. -->
# pivotTableAnchorReportType

## PivotTableAnchorReport

Exact pivot-table anchor report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| topLeftAddress | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| locationRange | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "PivotTableAnchorReport",
  "summary" : "Exact pivot-table anchor report.",
  "fields" : [ {
    "name" : "topLeftAddress",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "locationRange",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
