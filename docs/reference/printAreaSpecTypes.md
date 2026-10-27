<!-- Generated from core metadata; do not edit. -->
# printAreaSpecTypes

## NONE

Sheet has no explicit print area.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Sheet has no explicit print area.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## RANGE

Sheet prints the provided rectangular A1-style range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "RANGE",
  "summary" : "Sheet prints the provided rectangular A1-style range.",
  "fields" : [ {
    "name" : "range",
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
