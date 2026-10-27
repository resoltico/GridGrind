<!-- Generated from core metadata; do not edit. -->
# printTitleRowsSpecTypes

## NONE

Sheet has no repeating print-title rows.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Sheet has no repeating print-title rows.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## BAND

Sheet repeats the provided inclusive zero-based row band on every printed page.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| firstRowIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| lastRowIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "BAND",
  "summary" : "Sheet repeats the provided inclusive zero-based row band on every printed page.",
  "fields" : [ {
    "name" : "firstRowIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "lastRowIndex",
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
