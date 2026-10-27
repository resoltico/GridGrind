<!-- Generated from core metadata; do not edit. -->
# cellAlignmentReportType

## CellAlignmentReport

Exact cell-alignment report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| wrapText | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| horizontalAlignment | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| verticalAlignment | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| textRotation | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| indentation | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "CellAlignmentReport",
  "summary" : "Exact cell-alignment report.",
  "fields" : [ {
    "name" : "wrapText",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "horizontalAlignment",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "GENERAL", "LEFT", "CENTER", "RIGHT", "FILL", "JUSTIFY", "CENTER_SELECTION", "DISTRIBUTED" ]
  }, {
    "name" : "verticalAlignment",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "TOP", "CENTER", "BOTTOM", "JUSTIFY", "DISTRIBUTED" ]
  }, {
    "name" : "textRotation",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "indentation",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
