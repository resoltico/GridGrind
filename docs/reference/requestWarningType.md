<!-- Generated from core metadata; do not edit. -->
# requestWarningType

## RequestWarning

Non-fatal authored-plan warning surfaced at the response root.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| code | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| location | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"requestWarningLocationTypes"} |  |
| message | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "RequestWarning",
  "summary" : "Non-fatal authored-plan warning surfaced at the response root.",
  "fields" : [ {
    "name" : "code",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "UNQUOTED_SHEET_NAME_IN_FORMULA", "NON_PORTABLE_ABSOLUTE_PATH", "UTF8_BOM_IGNORED", "FORMULA_NOT_EVALUATED" ]
  }, {
    "name" : "location",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "requestWarningLocationTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "message",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
