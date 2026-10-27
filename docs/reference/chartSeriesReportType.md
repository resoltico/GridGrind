<!-- Generated from core metadata; do not edit. -->
# chartSeriesReportType

## ChartSeriesReport

Exact chart-series report. smooth, marker, and explosion fields are populated only when the stored plot family supports them.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| title | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"titleTypes"} |  |
| categories | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"dataSourceTypes"} |  |
| values | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"dataSourceTypes"} |  |
| smooth | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| markerStyle | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| markerSize | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| explosion | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "ChartSeriesReport",
  "summary" : "Exact chart-series report. smooth, marker, and explosion fields are populated only when the stored plot family supports them.",
  "fields" : [ {
    "name" : "title",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "titleTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "categories",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "dataSourceTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "values",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "dataSourceTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "smooth",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "markerStyle",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "CIRCLE", "DASH", "DIAMOND", "DOT", "NONE", "PICTURE", "PLUS", "SQUARE", "STAR", "TRIANGLE", "X" ]
  }, {
    "name" : "markerSize",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "explosion",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
