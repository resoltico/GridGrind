<!-- Generated from core metadata; do not edit. -->
# autofilterCustomConditionInputType

## AutofilterCustomConditionInput

One comparator-value pair nested inside a custom autofilter criterion.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| operator | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| value | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "AutofilterCustomConditionInput",
  "summary" : "One comparator-value pair nested inside a custom autofilter criterion.",
  "fields" : [ {
    "name" : "operator",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "value",
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
