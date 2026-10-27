<!-- Generated from core metadata; do not edit. -->
# formulaUdfFunctionSpecType

## FormulaUdfFunctionSpec

One template-backed user-defined function. formulaTemplate may reference ARG1, ARG2, and higher placeholders. maximumArgumentCount defaults to minimumArgumentCount when omitted.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| minimumArgumentCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| maximumArgumentCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| formulaTemplate | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "FormulaUdfFunctionSpec",
  "summary" : "One template-backed user-defined function. formulaTemplate may reference ARG1, ARG2, and higher placeholders. maximumArgumentCount defaults to minimumArgumentCount when omitted.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "minimumArgumentCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "maximumArgumentCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "formulaTemplate",
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
