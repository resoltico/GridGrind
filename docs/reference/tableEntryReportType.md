<!-- Generated from core metadata; do not edit. -->
# tableEntryReportType

## TableEntryReport

Exact workbook table report used by table-facts assertions.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| structure | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"tableEntryStructureReportType"} |  |
| style | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"tableStyleReportTypes"} |  |
| behavior | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"tableEntryBehaviorReportType"} | {hasAutofilter=false, published=false, insertRow=false, insertRowShift=false} |
| presentation | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"tableEntryPresentationReportType"} | {} |

```json
{
  "id" : "TableEntryReport",
  "summary" : "Exact workbook table report used by table-facts assertions.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "sheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "range",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "structure",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "tableEntryStructureReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "style",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "tableStyleReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "behavior",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "tableEntryBehaviorReportType"
    },
    "enumValues" : [ ],
    "defaultValue" : {
      "hasAutofilter" : false,
      "published" : false,
      "insertRow" : false,
      "insertRowShift" : false
    }
  }, {
    "name" : "presentation",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "tableEntryPresentationReportType"
    },
    "enumValues" : [ ],
    "defaultValue" : { }
  } ],
  "targetSelectors" : [ ]
}
```
