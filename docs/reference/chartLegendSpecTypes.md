<!-- Generated from core metadata; do not edit. -->
# chartLegendSpecTypes

## HIDDEN

Hide the legend entirely.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "HIDDEN",
  "summary" : "Hide the legend entirely.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## VISIBLE

Show the legend at one explicit position.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| position | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "VISIBLE",
  "summary" : "Show the legend at one explicit position.",
  "fields" : [ {
    "name" : "position",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "BOTTOM", "LEFT", "RIGHT", "TOP", "TOP_RIGHT" ]
  } ],
  "targetSelectors" : [ ]
}
```
