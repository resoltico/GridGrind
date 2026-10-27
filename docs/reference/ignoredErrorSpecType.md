<!-- Generated from core metadata; do not edit. -->
# ignoredErrorSpecType

## IgnoredErrorSpec

One ignored-error block anchored to one A1-style range plus one or more ignored-error families.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| errorTypes | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "IgnoredErrorSpec",
  "summary" : "One ignored-error block anchored to one A1-style range plus one or more ignored-error families.",
  "fields" : [ {
    "name" : "range",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "errorTypes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "STRING"
      }
    },
    "enumValues" : [ "CALCULATED_COLUMN", "EMPTY_CELL_REFERENCE", "EVALUATION_ERROR", "FORMULA", "FORMULA_RANGE", "LIST_DATA_VALIDATION", "NUMBER_STORED_AS_TEXT", "TWO_DIGIT_TEXT_YEAR", "UNLOCKED_FORMULA" ]
  } ],
  "targetSelectors" : [ ]
}
```
