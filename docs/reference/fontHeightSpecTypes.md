<!-- Generated from core metadata; do not edit. -->
# fontHeightSpecTypes

## POINTS

Specify font height in point units. Write format: {"type":"POINTS","points":13}. Read-back (GET_CELLS, GET_WINDOW): style.fontHeight is {"twips":260,"points":13} with both fields present, not this discriminated type format.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| points | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "POINTS",
  "summary" : "Specify font height in point units. Write format: {\"type\":\"POINTS\",\"points\":13}. Read-back (GET_CELLS, GET_WINDOW): style.fontHeight is {\"twips\":260,\"points\":13} with both fields present, not this discriminated type format.",
  "fields" : [ {
    "name" : "points",
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

## TWIPS

Specify font height in exact twips (20 twips = 1 point). Write format: {"type":"TWIPS","twips":260}. Read-back returns the same plain object shape as POINTS.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| twips | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "TWIPS",
  "summary" : "Specify font height in exact twips (20 twips = 1 point). Write format: {\"type\":\"TWIPS\",\"twips\":260}. Read-back returns the same plain object shape as POINTS.",
  "fields" : [ {
    "name" : "twips",
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
