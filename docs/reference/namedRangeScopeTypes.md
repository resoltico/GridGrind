<!-- Generated from core metadata; do not edit. -->
# namedRangeScopeTypes

## WORKBOOK

Target workbook scope.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "WORKBOOK",
  "summary" : "Target workbook scope.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## SHEET

Target one specific sheet scope.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "SHEET",
  "summary" : "Target one specific sheet scope.",
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
