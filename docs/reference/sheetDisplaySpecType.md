<!-- Generated from core metadata; do not edit. -->
# sheetDisplaySpecType

## SheetDisplaySpec

Screen-facing sheet display flags authored as part of sheet-presentation state. Supply all display flags explicitly.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| displayGridlines | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| displayZeros | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| displayRowColHeadings | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| displayFormulas | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| rightToLeft | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "SheetDisplaySpec",
  "summary" : "Screen-facing sheet display flags authored as part of sheet-presentation state. Supply all display flags explicitly.",
  "fields" : [ {
    "name" : "displayGridlines",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "displayZeros",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "displayRowColHeadings",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "displayFormulas",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "rightToLeft",
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
