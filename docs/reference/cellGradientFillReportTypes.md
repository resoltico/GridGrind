<!-- Generated from core metadata; do not edit. -->
# cellGradientFillReportTypes

## LINEAR

Read one factual linear gradient fill with ordered stops.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| degree | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| stops | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"cellGradientStopReportType"}} |  |

```json
{
  "id" : "LINEAR",
  "summary" : "Read one factual linear gradient fill with ordered stops.",
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
        "group" : "cellGradientStopReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## PATH

Read one factual path gradient fill with ordered stops. Edge offsets are omitted when Excel does not persist them.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| left | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| right | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| top | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| bottom | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| stops | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"cellGradientStopReportType"}} |  |

```json
{
  "id" : "PATH",
  "summary" : "Read one factual path gradient fill with ordered stops. Edge offsets are omitted when Excel does not persist them.",
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
        "group" : "cellGradientStopReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
