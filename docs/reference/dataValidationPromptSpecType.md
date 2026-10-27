<!-- Generated from core metadata; do not edit. -->
# dataValidationPromptSpecType

## DataValidationPromptSpec

Optional prompt-box configuration shown when a validated cell is selected.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| title | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |
| text | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |
| showPromptBox | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | true |

```json
{
  "id" : "DataValidationPromptSpec",
  "summary" : "Optional prompt-box configuration shown when a validated cell is selected.",
  "fields" : [ {
    "name" : "title",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "text",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "showPromptBox",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : true
  } ],
  "targetSelectors" : [ ]
}
```
