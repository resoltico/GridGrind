<!-- Generated from core metadata; do not edit. -->
# cellFontPatchType

## CellFontPatch

Font patch for cell styling; at least one field must be set. Colors can use RGB, theme, indexed, and tint semantics.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| bold | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| italic | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| fontName | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| fontHeight | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"fontHeightSpecTypes"} |  |
| fontColor | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |
| underline | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| strikeout | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "CellFontPatch",
  "summary" : "Font patch for cell styling; at least one field must be set. Colors can use RGB, theme, indexed, and tint semantics.",
  "fields" : [ {
    "name" : "bold",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "italic",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "fontName",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "fontHeight",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "fontHeightSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "fontColor",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "underline",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "strikeout",
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
