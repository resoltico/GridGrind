<!-- Generated from core metadata; do not edit. -->
# cellAlignmentPatchType

## CellAlignmentPatch

Alignment patch for cell styling; at least one field must be set. textRotation uses XSSF's explicit 0-180 degree scale and indentation uses Excel's 0-250 cell-indent range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| wrapText | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| horizontalAlignment | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| verticalAlignment | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| textRotation | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| indentation | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "CellAlignmentPatch",
  "summary" : "Alignment patch for cell styling; at least one field must be set. textRotation uses XSSF's explicit 0-180 degree scale and indentation uses Excel's 0-250 cell-indent range.",
  "fields" : [ {
    "name" : "wrapText",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "horizontalAlignment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "GENERAL", "LEFT", "CENTER", "RIGHT", "FILL", "JUSTIFY", "CENTER_SELECTION", "DISTRIBUTED" ]
  }, {
    "name" : "verticalAlignment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "TOP", "CENTER", "BOTTOM", "JUSTIFY", "DISTRIBUTED" ]
  }, {
    "name" : "textRotation",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "indentation",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
