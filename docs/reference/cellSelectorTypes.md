<!-- Generated from core metadata; do not edit. -->
# cellSelectorTypes

## CELL_ALL_USED_IN_SHEET

Select every physically present cell on one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "CELL_ALL_USED_IN_SHEET",
  "summary" : "Select every physically present cell on one sheet.",
  "fields" : [ {
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

## CELL_BY_ADDRESS

Select one exact cell on one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| address | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "CELL_BY_ADDRESS",
  "summary" : "Select one exact cell on one sheet.",
  "fields" : [ {
    "name" : "sheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "address",
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

## CELL_BY_ADDRESSES

Select one or more exact cells on one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| addresses | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "CELL_BY_ADDRESSES",
  "summary" : "Select one or more exact cells on one sheet.",
  "fields" : [ {
    "name" : "sheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "addresses",
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
