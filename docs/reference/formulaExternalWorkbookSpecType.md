<!-- Generated from core metadata; do not edit. -->
# formulaExternalWorkbookSpecType

## FormulaExternalWorkbookSpec

One external workbook binding keyed by the workbook name used inside formulas. Workbook paths resolve through the shared request-owned path rule.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| workbookName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| path | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "FormulaExternalWorkbookSpec",
  "summary" : "One external workbook binding keyed by the workbook name used inside formulas. Workbook paths resolve through the shared request-owned path rule.",
  "fields" : [ {
    "name" : "workbookName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "path",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ],
  "noteRefs" : [ "requestOwnedPathRule" ]
}
```
