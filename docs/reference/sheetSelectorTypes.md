<!-- Generated from core metadata; do not edit. -->
# sheetSelectorTypes

## SHEET_ALL

Select every sheet in workbook order.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "SHEET_ALL",
  "summary" : "Select every sheet in workbook order.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## SHEET_BY_NAME

Select one exact sheet by name.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "SHEET_BY_NAME",
  "summary" : "Select one exact sheet by name.",
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

## SHEET_BY_NAMES

Select one or more exact sheets by ordered name list.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| names | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "SHEET_BY_NAMES",
  "summary" : "Select one or more exact sheets by ordered name list.",
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
