<!-- Generated from core metadata; do not edit. -->
# cellStyleReportType

## CellStyleReport

Exact effective cell-style report used by style assertions.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| numberFormat | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| alignment | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"cellAlignmentReportType"} |  |
| font | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"cellFontReportType"} |  |
| fill | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellFillReportTypes"} |  |
| border | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"cellBorderReportType"} |  |
| protection | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"cellProtectionReportType"} |  |

```json
{
  "id" : "CellStyleReport",
  "summary" : "Exact effective cell-style report used by style assertions.",
  "fields" : [ {
    "name" : "numberFormat",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "alignment",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellAlignmentReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "font",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellFontReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "fill",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellFillReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "border",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellBorderReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "protection",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellProtectionReportType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
