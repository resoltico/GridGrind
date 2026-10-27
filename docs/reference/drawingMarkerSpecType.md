<!-- Generated from core metadata; do not edit. -->
# drawingMarkerSpecType

## DrawingMarkerSpec

One zero-based drawing marker with explicit column, row, and in-cell offsets.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| columnIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| rowIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| dx | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| dy | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "DrawingMarkerSpec",
  "summary" : "One zero-based drawing marker with explicit column, row, and in-cell offsets.",
  "fields" : [ {
    "name" : "columnIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "rowIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "dx",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "dy",
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
