<!-- Generated from core metadata; do not edit. -->
# cellGradientStopSpecType

## CellGradientStopSpec

One gradient stop with a normalized position between 0.0 and 1.0.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| position | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| color | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |

```json
{
  "id" : "CellGradientStopSpec",
  "summary" : "One gradient stop with a normalized position between 0.0 and 1.0.",
  "fields" : [ {
    "name" : "position",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "color",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
