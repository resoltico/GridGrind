<!-- Generated from core metadata; do not edit. -->
# executionPolicySpecType

## ExecutionPolicySpec

Optional request execution policy covering execution.mode, execution.journal, execution.calculation, and execution.assertionMode. Omit execution when the standard FULL_XSSF / SUMMARY / DO_NOT_CALCULATE / FAIL_FAST policy is intended, and omit any nested execution field that should keep its own default.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| mode | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"executionModeSpecTypes"} | {type=FULL_XSSF} |
| journal | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"executionJournalSpecType"} | {level=SUMMARY} |
| calculation | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"calculationPolicySpecType"} | {strategy={type=DO_NOT_CALCULATE}, markRecalculateOnOpen=false} |
| assertionMode | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} | FAIL_FAST |

```json
{
  "id" : "ExecutionPolicySpec",
  "summary" : "Optional request execution policy covering execution.mode, execution.journal, execution.calculation, and execution.assertionMode. Omit execution when the standard FULL_XSSF / SUMMARY / DO_NOT_CALCULATE / FAIL_FAST policy is intended, and omit any nested execution field that should keep its own default.",
  "fields" : [ {
    "name" : "mode",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "executionModeSpecTypes"
    },
    "enumValues" : [ ],
    "defaultValue" : {
      "type" : "FULL_XSSF"
    }
  }, {
    "name" : "journal",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "executionJournalSpecType"
    },
    "enumValues" : [ ],
    "defaultValue" : {
      "level" : "SUMMARY"
    }
  }, {
    "name" : "calculation",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "calculationPolicySpecType"
    },
    "enumValues" : [ ],
    "defaultValue" : {
      "strategy" : {
        "type" : "DO_NOT_CALCULATE"
      },
      "markRecalculateOnOpen" : false
    }
  }, {
    "name" : "assertionMode",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "FAIL_FAST", "COLLECT" ],
    "defaultValue" : "FAIL_FAST"
  } ],
  "targetSelectors" : [ ]
}
```
