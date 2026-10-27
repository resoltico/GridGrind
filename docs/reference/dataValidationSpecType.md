<!-- Generated from core metadata; do not edit. -->
# dataValidationSpecType

## DataValidationSpec

Supported data-validation definition attached to one sheet range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| rule | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"dataValidationRuleSpecTypes"} |  |
| allowBlank | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| suppressDropDownArrow | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| prompt | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"dataValidationPromptSpecType"} |  |
| errorAlert | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"dataValidationErrorAlertSpecType"} |  |

```json
{
  "id" : "DataValidationSpec",
  "summary" : "Supported data-validation definition attached to one sheet range.",
  "fields" : [ {
    "name" : "rule",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "dataValidationRuleSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "allowBlank",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "suppressDropDownArrow",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "prompt",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "dataValidationPromptSpecType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "errorAlert",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "dataValidationErrorAlertSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
