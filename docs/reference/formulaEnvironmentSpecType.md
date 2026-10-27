<!-- Generated from core metadata; do not edit. -->
# formulaEnvironmentSpecType

## FormulaEnvironmentSpec

Optional request-scoped formula-evaluation environment covering external workbook bindings, missing-workbook policy, and template-backed UDF toolpacks. Omit formulaEnvironment when the default evaluator is intended.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| externalWorkbooks | OPTIONAL | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"formulaExternalWorkbookSpecType"}} | [] |
| missingWorkbookPolicy | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} | ERROR |
| udfToolpacks | OPTIONAL | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"formulaUdfToolpackSpecType"}} | [] |

```json
{
  "id" : "FormulaEnvironmentSpec",
  "summary" : "Optional request-scoped formula-evaluation environment covering external workbook bindings, missing-workbook policy, and template-backed UDF toolpacks. Omit formulaEnvironment when the default evaluator is intended.",
  "fields" : [ {
    "name" : "externalWorkbooks",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "formulaExternalWorkbookSpecType"
      }
    },
    "enumValues" : [ ],
    "defaultValue" : [ ]
  }, {
    "name" : "missingWorkbookPolicy",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "ERROR", "USE_CACHED_VALUE" ],
    "defaultValue" : "ERROR"
  }, {
    "name" : "udfToolpacks",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "formulaUdfToolpackSpecType"
      }
    },
    "enumValues" : [ ],
    "defaultValue" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
