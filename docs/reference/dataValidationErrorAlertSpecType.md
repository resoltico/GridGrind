<!-- Generated from core metadata; do not edit. -->
# dataValidationErrorAlertSpecType

## DataValidationErrorAlertSpec

Optional error-box configuration shown when invalid data is entered.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| style | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| title | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |
| text | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |
| showErrorBox | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | true |

```json
{
  "id" : "DataValidationErrorAlertSpec",
  "summary" : "Optional error-box configuration shown when invalid data is entered.",
  "fields" : [ {
    "name" : "style",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "STOP", "WARNING", "INFORMATION" ]
  }, {
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
    "name" : "showErrorBox",
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
