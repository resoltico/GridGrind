<!-- Generated from core metadata; do not edit. -->
# pivotTableFieldReportType

## PivotTableFieldReport

Exact pivot field report bound to one source column.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sourceColumnIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| sourceColumnName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "PivotTableFieldReport",
  "summary" : "Exact pivot field report bound to one source column.",
  "fields" : [ {
    "name" : "sourceColumnIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "sourceColumnName",
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
