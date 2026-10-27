<!-- Generated from core metadata; do not edit. -->
# cellStylePatchType

## CellStylePatch

Style patch applied to a cell or range; at least one field must be set. Colors preserve RGB, theme, indexed, and tint semantics; style subgroups are nested explicitly.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| numberFormat | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| alignment | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellAlignmentPatchType"} |  |
| font | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellFontPatchType"} |  |
| fill | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"cellFillSpecTypes"} |  |
| border | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellBorderPatchType"} |  |
| protection | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellProtectionPatchType"} |  |

```json
{
  "id" : "CellStylePatch",
  "summary" : "Style patch applied to a cell or range; at least one field must be set. Colors preserve RGB, theme, indexed, and tint semantics; style subgroups are nested explicitly.",
  "fields" : [ {
    "name" : "numberFormat",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "alignment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellAlignmentPatchType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "font",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellFontPatchType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "fill",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellFillSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "border",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellBorderPatchType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "protection",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellProtectionPatchType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
