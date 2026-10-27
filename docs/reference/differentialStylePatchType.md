<!-- Generated from core metadata; do not edit. -->
# differentialStylePatchType

## DifferentialStylePatch

Differential style payload used by authored conditional-formatting rules. At least one field must be set. Colors preserve RGB, theme, indexed, and tint semantics.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| numberFormat | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| bold | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| italic | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| fontHeight | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"fontHeightSpecTypes"} |  |
| fontColor | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |
| underline | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| strikeout | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| fillColor | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |
| border | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"differentialBorderPatchType"} |  |

```json
{
  "id" : "DifferentialStylePatch",
  "summary" : "Differential style payload used by authored conditional-formatting rules. At least one field must be set. Colors preserve RGB, theme, indexed, and tint semantics.",
  "fields" : [ {
    "name" : "numberFormat",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
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
  }, {
    "name" : "fillColor",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "border",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "differentialBorderPatchType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
