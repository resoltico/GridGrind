<!-- Generated from core metadata; do not edit. -->
# namedRangeScopeReportTypes

## WORKBOOK

Observed workbook scope.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "WORKBOOK",
  "summary" : "Observed workbook scope.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## SHEET

Observed local sheet scope.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "SHEET",
  "summary" : "Observed local sheet scope.",
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
