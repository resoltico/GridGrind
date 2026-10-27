<!-- Generated from core metadata; do not edit. -->
# pivotTableReportTypes

## SUPPORTED

Exact supported pivot-table report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| anchor | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"pivotTableAnchorReportType"} |  |
| source | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"pivotSourceReportTypes"} |  |
| rowLabels | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"pivotTableFieldReportType"}} |  |
| columnLabels | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"pivotTableFieldReportType"}} |  |
| reportFilters | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"pivotTableFieldReportType"}} |  |
| dataFields | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"pivotTableDataFieldReportType"}} |  |
| valuesAxisOnColumns | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "SUPPORTED",
  "summary" : "Exact supported pivot-table report.",
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
    "name" : "anchor",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "pivotTableAnchorReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "source",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "pivotSourceReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "rowLabels",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "pivotTableFieldReportType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "columnLabels",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "pivotTableFieldReportType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "reportFilters",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "pivotTableFieldReportType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "dataFields",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "pivotTableDataFieldReportType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "valuesAxisOnColumns",
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

## UNSUPPORTED

Exact unsupported pivot-table report preserved from the workbook.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| anchor | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"pivotTableAnchorReportType"} |  |
| detail | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "UNSUPPORTED",
  "summary" : "Exact unsupported pivot-table report preserved from the workbook.",
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
    "name" : "anchor",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "pivotTableAnchorReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "detail",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
