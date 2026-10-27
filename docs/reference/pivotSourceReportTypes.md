<!-- Generated from core metadata; do not edit. -->
# pivotSourceReportTypes

## RANGE

Pivot source resolved from one sheet range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "RANGE",
  "summary" : "Pivot source resolved from one sheet range.",
  "fields" : [ {
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
  } ],
  "targetSelectors" : [ ]
}
```

## NAMED_RANGE

Pivot source resolved from one named range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "NAMED_RANGE",
  "summary" : "Pivot source resolved from one named range.",
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
  } ],
  "targetSelectors" : [ ]
}
```

## TABLE

Pivot source resolved from one workbook table.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "TABLE",
  "summary" : "Pivot source resolved from one workbook table.",
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
  } ],
  "targetSelectors" : [ ]
}
```
