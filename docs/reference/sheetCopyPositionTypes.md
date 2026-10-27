<!-- Generated from core metadata; do not edit. -->
# sheetCopyPositionTypes

## APPEND_AT_END

Place the copied sheet after every existing sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "APPEND_AT_END",
  "summary" : "Place the copied sheet after every existing sheet.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## AT_INDEX

Place the copied sheet at the requested zero-based workbook position.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| targetIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "AT_INDEX",
  "summary" : "Place the copied sheet at the requested zero-based workbook position.",
  "fields" : [ {
    "name" : "targetIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
