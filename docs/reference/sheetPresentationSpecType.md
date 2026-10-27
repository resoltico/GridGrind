<!-- Generated from core metadata; do not edit. -->
# sheetPresentationSpecType

## SheetPresentationSpec

Authoritative sheet-presentation payload for one SET_SHEET_PRESENTATION request. Supply explicit display, tabColor, outlineSummary, sheetDefaults, and ignoredErrors values.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| display | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"sheetDisplaySpecType"} |  |
| tabColor | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |
| outlineSummary | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"sheetOutlineSummarySpecType"} |  |
| sheetDefaults | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"sheetDefaultsSpecType"} |  |
| ignoredErrors | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"ignoredErrorSpecType"}} |  |

```json
{
  "id" : "SheetPresentationSpec",
  "summary" : "Authoritative sheet-presentation payload for one SET_SHEET_PRESENTATION request. Supply explicit display, tabColor, outlineSummary, sheetDefaults, and ignoredErrors values.",
  "fields" : [ {
    "name" : "display",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "sheetDisplaySpecType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "tabColor",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "outlineSummary",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "sheetOutlineSummarySpecType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "sheetDefaults",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "sheetDefaultsSpecType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "ignoredErrors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "ignoredErrorSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
