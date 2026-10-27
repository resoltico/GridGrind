<!-- Generated from core metadata; do not edit. -->
# dataValidationPromptReportType

## DataValidationPromptReport

Observed prompt-box text and display state.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| title | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| text | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| showPromptBox | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "DataValidationPromptReport",
  "summary" : "Observed prompt-box text and display state.",
  "fields" : [ {
    "name" : "title",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "text",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "showPromptBox",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
