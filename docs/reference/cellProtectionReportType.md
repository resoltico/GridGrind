<!-- Generated from core metadata; do not edit. -->
# cellProtectionReportType

## CellProtectionReport

Exact cell-protection report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| locked | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| hiddenFormula | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "CellProtectionReport",
  "summary" : "Exact cell-protection report.",
  "fields" : [ {
    "name" : "locked",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "hiddenFormula",
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
