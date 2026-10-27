<!-- Generated from core metadata; do not edit. -->
# sheetSummaryReportType

## SheetSummaryReport

Exact sheet summary report including visibility, protection, and structural counts.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| visibility | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| protection | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"sheetProtectionReportTypes"} |  |
| physicalRowCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| lastRowIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| lastColumnIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "SheetSummaryReport",
  "summary" : "Exact sheet summary report including visibility, protection, and structural counts.",
  "fields" : [ {
    "name" : "sheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "visibility",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "VISIBLE", "HIDDEN", "VERY_HIDDEN" ]
  }, {
    "name" : "protection",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "sheetProtectionReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "physicalRowCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "lastRowIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "lastColumnIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
