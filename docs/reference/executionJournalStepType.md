<!-- Generated from core metadata; do not edit. -->
# executionJournalStepType

## ExecutionJournalStep

Per-step execution telemetry with resolved targets, timing, outcome, and optional failure detail.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| stepIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| stepId | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| stepKind | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| stepType | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| resolvedTargets | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"executionJournalTargetType"}} |  |
| phase | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"phaseTypes"} |  |
| outcome | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| failure | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"executionJournalFailureClassificationType"} |  |

```json
{
  "id" : "ExecutionJournalStep",
  "summary" : "Per-step execution telemetry with resolved targets, timing, outcome, and optional failure detail.",
  "fields" : [ {
    "name" : "stepIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "stepId",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "stepKind",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "stepType",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "resolvedTargets",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "executionJournalTargetType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "phase",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "phaseTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "outcome",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "SUCCEEDED", "FAILED" ]
  }, {
    "name" : "failure",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "executionJournalFailureClassificationType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
