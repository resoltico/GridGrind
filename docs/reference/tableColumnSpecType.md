<!-- Generated from core metadata; do not edit. -->
# tableColumnSpecType

## TableColumnSpec

Advanced table-column metadata applied by zero-based ordinal column index.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| columnIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| uniqueName | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| totalsRowLabel | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| totalsRowFunction | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| calculatedColumnFormula | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "TableColumnSpec",
  "summary" : "Advanced table-column metadata applied by zero-based ordinal column index.",
  "fields" : [ {
    "name" : "columnIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "uniqueName",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "defaultValue" : ""
  }, {
    "name" : "totalsRowLabel",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "defaultValue" : ""
  }, {
    "name" : "totalsRowFunction",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "defaultValue" : ""
  }, {
    "name" : "calculatedColumnFormula",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "defaultValue" : ""
  } ],
  "targetSelectors" : [ ]
}
```
