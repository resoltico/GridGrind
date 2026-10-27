<!-- Generated from core metadata; do not edit. -->
# pivotTableDataFieldInputType

## PivotTableDataFieldInput

One authored pivot data field bound to a source column and aggregation function. displayName defaults to sourceColumnName when omitted.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sourceColumnName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| function | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| displayName | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} | from sourceColumnName |
| valueFormat | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "PivotTableDataFieldInput",
  "summary" : "One authored pivot data field bound to a source column and aggregation function. displayName defaults to sourceColumnName when omitted.",
  "fields" : [ {
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
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "defaultFrom" : "sourceColumnName"
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
