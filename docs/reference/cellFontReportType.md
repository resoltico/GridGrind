<!-- Generated from core metadata; do not edit. -->
# cellFontReportType

## CellFontReport

Exact cell-font report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| bold | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| italic | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| fontName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| fontHeight | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"fontHeightReportType"} |  |
| fontColor | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"cellColorReportTypes"} |  |
| underline | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| strikeout | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "CellFontReport",
  "summary" : "Exact cell-font report.",
  "fields" : [ {
    "name" : "bold",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "italic",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "fontName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "fontHeight",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "fontHeightReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "fontColor",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellColorReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "underline",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "strikeout",
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
