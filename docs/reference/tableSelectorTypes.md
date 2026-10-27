<!-- Generated from core metadata; do not edit. -->
# tableSelectorTypes

## TABLE_ALL

Select every table in workbook order.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "TABLE_ALL",
  "summary" : "Select every table in workbook order.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## TABLE_BY_NAME

Select one workbook-global table by exact name.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "TABLE_BY_NAME",
  "summary" : "Select one workbook-global table by exact name.",
  "fields" : [ {
    "name" : "name",
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

## TABLE_BY_NAMES

Select one or more workbook-global tables by exact name.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| names | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "TABLE_BY_NAMES",
  "summary" : "Select one or more workbook-global tables by exact name.",
  "fields" : [ {
    "name" : "names",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "STRING"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## TABLE_BY_NAME_ON_SHEET

Select one workbook-global table by exact name and expected owning sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "TABLE_BY_NAME_ON_SHEET",
  "summary" : "Select one workbook-global table by exact name and expected owning sheet.",
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
  } ],
  "targetSelectors" : [ ]
}
```
