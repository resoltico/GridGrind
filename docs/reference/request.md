<!-- Generated from core metadata; do not edit. -->
# Request

## WorkbookPlan

Complete GridGrind plan for workbook source, execution policy, formula environment, ordered mutation, assertion, and inspection steps, and persistence.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| protocolVersion | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":1.0,"maximum":1.0}]} |  |
| planId | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| source | REQUIRED | {"type":"TOP_LEVEL_TYPE_SET","typeSet":"workbookSourceTypes"} |  |
| persistence | REQUIRED | {"type":"TOP_LEVEL_TYPE_SET","typeSet":"workbookPersistenceTypes"} |  |
| execution | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"executionPolicySpecType"} | {mode={type=FULL_XSSF}, journal={level=SUMMARY}, calculation={strategy={type=DO_NOT_CALCULATE}, markRecalculateOnOpen=false}, assertionMode=FAIL_FAST} |
| formulaEnvironment | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"formulaEnvironmentSpecType"} | {externalWorkbooks=[], missingWorkbookPolicy=ERROR, udfToolpacks=[]} |
| steps | REQUIRED | {"type":"LIST","elementShape":{"type":"TOP_LEVEL_TYPE_SET","typeSet":"workbookStepTypes"}} |  |

```json
{
  "id" : "WorkbookPlan",
  "summary" : "Complete GridGrind plan for workbook source, execution policy, formula environment, ordered mutation, assertion, and inspection steps, and persistence.",
  "fields" : [ {
    "name" : "protocolVersion",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "NUMBER_RANGE",
        "minimum" : 1.0,
        "maximum" : 1.0
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "planId",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "source",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "TOP_LEVEL_TYPE_SET",
      "typeSet" : "workbookSourceTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "persistence",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "TOP_LEVEL_TYPE_SET",
      "typeSet" : "workbookPersistenceTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "execution",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "executionPolicySpecType"
    },
    "enumValues" : [ ],
    "defaultValue" : {
      "mode" : {
        "type" : "FULL_XSSF"
      },
      "journal" : {
        "level" : "SUMMARY"
      },
      "calculation" : {
        "strategy" : {
          "type" : "DO_NOT_CALCULATE"
        },
        "markRecalculateOnOpen" : false
      },
      "assertionMode" : "FAIL_FAST"
    }
  }, {
    "name" : "formulaEnvironment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "formulaEnvironmentSpecType"
    },
    "enumValues" : [ ],
    "defaultValue" : {
      "externalWorkbooks" : [ ],
      "missingWorkbookPolicy" : "ERROR",
      "udfToolpacks" : [ ]
    }
  }, {
    "name" : "steps",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "TOP_LEVEL_TYPE_SET",
        "typeSet" : "workbookStepTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
