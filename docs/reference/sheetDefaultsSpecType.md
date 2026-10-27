<!-- Generated from core metadata; do not edit. -->
# sheetDefaultsSpecType

## SheetDefaultsSpec

Default row and column sizing authored as part of sheet-presentation state. Supply explicit defaultColumnWidth and defaultRowHeightPoints values. defaultColumnWidth must be > 0 and <= 255; defaultRowHeightPoints must be > 0 and <= 409.0.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| defaultColumnWidth | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| defaultRowHeightPoints | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "SheetDefaultsSpec",
  "summary" : "Default row and column sizing authored as part of sheet-presentation state. Supply explicit defaultColumnWidth and defaultRowHeightPoints values. defaultColumnWidth must be > 0 and <= 255; defaultRowHeightPoints must be > 0 and <= 409.0.",
  "fields" : [ {
    "name" : "defaultColumnWidth",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "defaultRowHeightPoints",
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
