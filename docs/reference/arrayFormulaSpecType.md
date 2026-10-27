<!-- Generated from core metadata; do not edit. -->
# arrayFormulaSpecType

## ArrayFormulaSpec

One authored array formula bound to a contiguous single-cell or multi-cell range. Leading = or {=...} wrappers normalize away for inline sources.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| source | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |

```json
{
  "id" : "ArrayFormulaSpec",
  "summary" : "One authored array formula bound to a contiguous single-cell or multi-cell range. Leading = or {=...} wrappers normalize away for inline sources.",
  "fields" : [ {
    "name" : "source",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
