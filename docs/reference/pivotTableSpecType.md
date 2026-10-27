<!-- Generated from core metadata; do not edit. -->
# pivotTableSpecType

## PivotTableSpec

Workbook-global pivot-table definition for one SET_PIVOT_TABLE request. PivotSourceSpec-column assignments across rowLabels, columnLabels, reportFilters, and dataFields must be disjoint. reportFilters require anchor.topLeftAddress on row 3 or lower.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| source | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"pivotSourceSpecTypes"} |  |
| anchor | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"pivotTableAnchorInputType"} |  |
| rowLabels | OPTIONAL | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} | [] |
| columnLabels | OPTIONAL | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} | [] |
| reportFilters | OPTIONAL | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} | [] |
| dataFields | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"pivotTableDataFieldInputType"}} |  |

```json
{
  "id" : "PivotTableSpec",
  "summary" : "Workbook-global pivot-table definition for one SET_PIVOT_TABLE request. PivotSourceSpec-column assignments across rowLabels, columnLabels, reportFilters, and dataFields must be disjoint. reportFilters require anchor.topLeftAddress on row 3 or lower.",
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
    "name" : "source",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "pivotSourceSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "anchor",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "pivotTableAnchorInputType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "rowLabels",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "STRING"
      }
    },
    "enumValues" : [ ],
    "defaultValue" : [ ]
  }, {
    "name" : "columnLabels",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "STRING"
      }
    },
    "enumValues" : [ ],
    "defaultValue" : [ ]
  }, {
    "name" : "reportFilters",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "STRING"
      }
    },
    "enumValues" : [ ],
    "defaultValue" : [ ]
  }, {
    "name" : "dataFields",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "pivotTableDataFieldInputType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
