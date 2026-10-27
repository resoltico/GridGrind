<!-- Generated from core metadata; do not edit. -->
# borderSidePatchType

## BorderSidePatch

One border side shared by cell and differential style patches, defined by its border style and optional color semantics.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| style | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| color | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |

```json
{
  "id" : "BorderSidePatch",
  "summary" : "One border side shared by cell and differential style patches, defined by its border style and optional color semantics.",
  "fields" : [ {
    "name" : "style",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "NONE", "THIN", "MEDIUM", "DASHED", "DOTTED", "THICK", "DOUBLE", "HAIR", "MEDIUM_DASHED", "DASH_DOT", "MEDIUM_DASH_DOT", "DASH_DOT_DOT", "MEDIUM_DASH_DOT_DOT", "SLANTED_DASH_DOT" ]
  }, {
    "name" : "color",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
