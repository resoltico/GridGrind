<!-- Generated from core metadata; do not edit. -->
# cellBorderReportType

## CellBorderReport

Exact four-sided cell-border report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| top | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellBorderSideReportTypes"} |  |
| right | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellBorderSideReportTypes"} |  |
| bottom | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellBorderSideReportTypes"} |  |
| left | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellBorderSideReportTypes"} |  |

```json
{
  "id" : "CellBorderReport",
  "summary" : "Exact four-sided cell-border report.",
  "fields" : [ {
    "name" : "top",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellBorderSideReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "right",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellBorderSideReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "bottom",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellBorderSideReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "left",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellBorderSideReportTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
