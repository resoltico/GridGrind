<!-- Generated from core metadata; do not edit. -->
# calculationPolicySpecType

## CalculationPolicySpec

Explicit formula-calculation policy. Use CalculationPolicySpec.defaults() for DO_NOT_CALCULATE with markRecalculateOnOpen=false. Use DEFERRED_CALCULATION to inspect capabilities without immediate evaluation, Use EVALUATE_ALL or EVALUATE_TARGETS for immediate server-side evaluation, CLEAR_CACHES_ONLY to strip persisted caches, or markRecalculateOnOpen=true when Excel-compatible clients should recalculate later.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| strategy | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"calculationStrategySpecTypes"} | {type=DO_NOT_CALCULATE} |
| markRecalculateOnOpen | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |

```json
{
  "id" : "CalculationPolicySpec",
  "summary" : "Explicit formula-calculation policy. Use CalculationPolicySpec.defaults() for DO_NOT_CALCULATE with markRecalculateOnOpen=false. Use DEFERRED_CALCULATION to inspect capabilities without immediate evaluation, Use EVALUATE_ALL or EVALUATE_TARGETS for immediate server-side evaluation, CLEAR_CACHES_ONLY to strip persisted caches, or markRecalculateOnOpen=true when Excel-compatible clients should recalculate later.",
  "fields" : [ {
    "name" : "strategy",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "calculationStrategySpecTypes"
    },
    "enumValues" : [ ],
    "defaultValue" : {
      "type" : "DO_NOT_CALCULATE"
    }
  }, {
    "name" : "markRecalculateOnOpen",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  } ],
  "targetSelectors" : [ ]
}
```
