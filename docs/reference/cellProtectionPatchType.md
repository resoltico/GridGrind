<!-- Generated from core metadata; do not edit. -->
# cellProtectionPatchType

## CellProtectionPatch

Cell protection patch; at least one field must be set. These flags matter when sheet protection is enabled.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| locked | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| hiddenFormula | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "CellProtectionPatch",
  "summary" : "Cell protection patch; at least one field must be set. These flags matter when sheet protection is enabled.",
  "fields" : [ {
    "name" : "locked",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "hiddenFormula",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
