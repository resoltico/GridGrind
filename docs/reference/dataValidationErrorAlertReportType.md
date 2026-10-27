<!-- Generated from core metadata; do not edit. -->
# dataValidationErrorAlertReportType

## DataValidationErrorAlertReport

Observed error-box text, style and display state.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| style | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| title | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| text | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| showErrorBox | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "DataValidationErrorAlertReport",
  "summary" : "Observed error-box text, style and display state.",
  "fields" : [ {
    "name" : "style",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "STOP", "WARNING", "INFORMATION" ]
  }, {
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
    "name" : "showErrorBox",
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
