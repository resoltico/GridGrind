<!-- Generated from core metadata; do not edit. -->
# executionJournalType

## ExecutionJournal

Structured execution telemetry returned on every success and failure response, including validation, input resolution, open, calculation, step, persistence, and close phases.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| level | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| source | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"executionJournalSourceSummaryType"} |  |
| validation | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"phaseTypes"} |  |
| inputResolution | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"phaseTypes"} |  |
| open | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"phaseTypes"} |  |
| calculation | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"executionJournalCalculationType"} |  |
| persistencePhase | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"phaseTypes"} |  |
| close | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"phaseTypes"} |  |
| steps | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"executionJournalStepType"}} |  |
| outcome | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"outcomeTypes"} |  |

```json
{
  "id" : "ExecutionJournal",
  "summary" : "Structured execution telemetry returned on every success and failure response, including validation, input resolution, open, calculation, step, persistence, and close phases.",
  "fields" : [ {
    "name" : "level",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "SUMMARY", "NORMAL", "VERBOSE" ]
  }, {
    "name" : "source",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "executionJournalSourceSummaryType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "validation",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "phaseTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "inputResolution",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "phaseTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "open",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "phaseTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "calculation",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "executionJournalCalculationType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "persistencePhase",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "phaseTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "close",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "phaseTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "steps",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "executionJournalStepType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "outcome",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "outcomeTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
