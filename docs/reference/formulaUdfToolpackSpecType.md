<!-- Generated from core metadata; do not edit. -->
# formulaUdfToolpackSpecType

## FormulaUdfToolpackSpec

One named collection of template-backed user-defined functions.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| functions | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"formulaUdfFunctionSpecType"}} |  |

```json
{
  "id" : "FormulaUdfToolpackSpec",
  "summary" : "One named collection of template-backed user-defined functions.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "functions",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "formulaUdfFunctionSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
