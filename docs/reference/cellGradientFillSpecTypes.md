<!-- Generated from core metadata; do not edit. -->
# cellGradientFillSpecTypes

## LINEAR

Write one linear gradient fill with ordered stops. degree is optional when Excel's default angle is acceptable.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| degree | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| stops | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"cellGradientStopSpecType"}} |  |

```json
{
  "id" : "LINEAR",
  "summary" : "Write one linear gradient fill with ordered stops. degree is optional when Excel's default angle is acceptable.",
  "fields" : [ {
    "name" : "degree",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "stops",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "cellGradientStopSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## PATH

Write one path gradient fill with ordered stops. Each edge offset is optional and omitted offsets preserve Excel's default path geometry.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| left | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| right | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| top | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| bottom | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| stops | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"cellGradientStopSpecType"}} |  |

```json
{
  "id" : "PATH",
  "summary" : "Write one path gradient fill with ordered stops. Each edge offset is optional and omitted offsets preserve Excel's default path geometry.",
  "fields" : [ {
    "name" : "left",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "right",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "top",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "bottom",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "stops",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "cellGradientStopSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
