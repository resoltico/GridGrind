<!-- Generated from core metadata; do not edit. -->
# tableColumnReportType

## TableColumnReport

Exact table-column report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| id | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| uniqueName | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| totalsRowLabel | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| totalsRowFunction | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| calculatedColumnFormula | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "TableColumnReport",
  "summary" : "Exact table-column report.",
  "fields" : [ {
    "name" : "id",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "uniqueName",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "totalsRowLabel",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "totalsRowFunction",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "calculatedColumnFormula",
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
