<!-- Generated from core metadata; do not edit. -->
# tableSpecType

## TableSpec

Workbook-global table definition for one SET_TABLE request.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"LENGTH_RANGE","minimum":1,"maximum":255},{"type":"STRING_PATTERN","pattern":"^[\\p{L}_\\\\][\\p{L}\\p{N}_.\\\\]*$"}]} |  |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| showTotalsRow | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| hasAutofilter | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | true |
| style | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"tableStyleSpecTypes"} |  |
| comment | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} | {type=INLINE, text=} |
| published | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| insertRow | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| insertRowShift | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| headerRowCellStyle | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| dataCellStyle | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| totalsRowCellStyle | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| columns | OPTIONAL | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"tableColumnSpecType"}} | [] |

```json
{
  "id" : "TableSpec",
  "summary" : "Workbook-global table definition for one SET_TABLE request.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "LENGTH_RANGE",
        "minimum" : 1,
        "maximum" : 255
      }, {
        "type" : "STRING_PATTERN",
        "pattern" : "^[\\p{L}_\\\\][\\p{L}\\p{N}_.\\\\]*$"
      } ]
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
    "name" : "showTotalsRow",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "hasAutofilter",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : true
  }, {
    "name" : "style",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "tableStyleSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "comment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ],
    "defaultValue" : {
      "type" : "INLINE",
      "text" : ""
    }
  }, {
    "name" : "published",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "insertRow",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "insertRowShift",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "headerRowCellStyle",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "defaultValue" : ""
  }, {
    "name" : "dataCellStyle",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "defaultValue" : ""
  }, {
    "name" : "totalsRowCellStyle",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "defaultValue" : ""
  }, {
    "name" : "columns",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "tableColumnSpecType"
      }
    },
    "enumValues" : [ ],
    "defaultValue" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
