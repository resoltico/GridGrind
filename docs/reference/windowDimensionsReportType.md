<!-- Generated from core metadata; do not edit. -->
# windowDimensionsReportType

## WindowDimensionsReport

Published row and column dimensions for one rectangular cell window.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| rowCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| columnCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "WindowDimensionsReport",
  "summary" : "Published row and column dimensions for one rectangular cell window.",
  "fields" : [ {
    "name" : "rowCount",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "columnCount",
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
