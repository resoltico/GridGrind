<!-- Generated from core metadata; do not edit. -->
# printLayoutSpecType

## PrintLayoutSpec

Authoritative supported print-layout payload for one SET_PRINT_LAYOUT request. Supply explicit printArea, orientation, scaling, repeatingRows, repeatingColumns, header, footer, and setup fields.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| printArea | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"printAreaSpecTypes"} |  |
| orientation | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| scaling | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"printScalingSpecTypes"} |  |
| repeatingRows | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"printTitleRowsSpecTypes"} |  |
| repeatingColumns | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"printTitleColumnsSpecTypes"} |  |
| header | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"headerFooterTextSpecType"} |  |
| footer | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"headerFooterTextSpecType"} |  |
| setup | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"printSetupSpecType"} |  |

```json
{
  "id" : "PrintLayoutSpec",
  "summary" : "Authoritative supported print-layout payload for one SET_PRINT_LAYOUT request. Supply explicit printArea, orientation, scaling, repeatingRows, repeatingColumns, header, footer, and setup fields.",
  "fields" : [ {
    "name" : "printArea",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "printAreaSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "orientation",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "PORTRAIT", "LANDSCAPE" ]
  }, {
    "name" : "scaling",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "printScalingSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "repeatingRows",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "printTitleRowsSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "repeatingColumns",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "printTitleColumnsSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "header",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "headerFooterTextSpecType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "footer",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "headerFooterTextSpecType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "setup",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "printSetupSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
