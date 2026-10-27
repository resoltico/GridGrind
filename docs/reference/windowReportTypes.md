<!-- Generated from core metadata; do not edit. -->
# windowReportTypes

## SPARSE

Sparse window omitting blank cells and publishing populatedCells only.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| topLeftAddress | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| dimensions | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"windowDimensionsReportType"} |  |
| populatedCells | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"cellReportTypes"}} |  |

```json
{
  "id" : "SPARSE",
  "summary" : "Sparse window omitting blank cells and publishing populatedCells only.",
  "fields" : [ {
    "name" : "sheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "topLeftAddress",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "dimensions",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "windowDimensionsReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "populatedCells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "cellReportTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## DENSE

Dense window preserving explicit row structure including blank cells.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| topLeftAddress | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| dimensions | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"windowDimensionsReportType"} |  |
| rows | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"windowRowReportType"}} |  |

```json
{
  "id" : "DENSE",
  "summary" : "Dense window preserving explicit row structure including blank cells.",
  "fields" : [ {
    "name" : "sheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "topLeftAddress",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "dimensions",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "windowDimensionsReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "rows",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "windowRowReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
