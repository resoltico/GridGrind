<!-- Generated from core metadata; do not edit. -->
# printTitleColumnsSpecTypes

## NONE

Sheet has no repeating print-title columns.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Sheet has no repeating print-title columns.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## BAND

Sheet repeats the provided inclusive zero-based column band on every printed page.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| firstColumnIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| lastColumnIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "BAND",
  "summary" : "Sheet repeats the provided inclusive zero-based column band on every printed page.",
  "fields" : [ {
    "name" : "firstColumnIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "lastColumnIndex",
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
