<!-- Generated from core metadata; do not edit. -->
# sheetOutlineSummarySpecType

## SheetOutlineSummarySpec

Outline-summary placement authored as part of sheet-presentation state. Supply both placement flags explicitly.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| rowSumsBelow | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| rowSumsRight | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "SheetOutlineSummarySpec",
  "summary" : "Outline-summary placement authored as part of sheet-presentation state. Supply both placement flags explicitly.",
  "fields" : [ {
    "name" : "rowSumsBelow",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "rowSumsRight",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
