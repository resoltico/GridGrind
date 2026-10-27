<!-- Generated from core metadata; do not edit. -->
# legendTypes

## HIDDEN

No legend is present.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "HIDDEN",
  "summary" : "No legend is present.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## VISIBLE

Visible legend at one persisted position.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| position | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "VISIBLE",
  "summary" : "Visible legend at one persisted position.",
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
