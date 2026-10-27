<!-- Generated from core metadata; do not edit. -->
# pivotTableDataFieldReportType

## PivotTableDataFieldReport

Exact pivot data-field report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sourceColumnIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| sourceColumnName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| function | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| displayName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| valueFormat | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "PivotTableDataFieldReport",
  "summary" : "Exact pivot data-field report.",
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
  }, {
    "name" : "function",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "SUM", "COUNT", "COUNT_NUMS", "AVERAGE", "MAX", "MIN", "PRODUCT", "STD_DEV", "STD_DEVP", "VAR", "VARP" ]
  }, {
    "name" : "displayName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "valueFormat",
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
