<!-- Generated from core metadata; do not edit. -->
# tableStyleSpecTypes

## NONE

Clear table style metadata.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Clear table style metadata.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## NAMED

Apply one named workbook table style with explicit stripe and emphasis flags.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| showFirstColumn | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| showLastColumn | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| showRowStripes | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| showColumnStripes | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "NAMED",
  "summary" : "Apply one named workbook table style with explicit stripe and emphasis flags.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "showFirstColumn",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "showLastColumn",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "showRowStripes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "showColumnStripes",
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
