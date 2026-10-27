<!-- Generated from core metadata; do not edit. -->
# tableEntryPresentationReportType

## TableEntryPresentationReport

Optional table comment and style labels attached to one persisted workbook table.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| comment | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| headerRowCellStyle | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| dataCellStyle | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| totalsRowCellStyle | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "TableEntryPresentationReport",
  "summary" : "Optional table comment and style labels attached to one persisted workbook table.",
  "fields" : [ {
    "name" : "comment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "headerRowCellStyle",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "dataCellStyle",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "totalsRowCellStyle",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
