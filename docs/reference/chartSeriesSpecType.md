<!-- Generated from core metadata; do not edit. -->
# chartSeriesSpecType

## ChartSeriesSpec

One authored chart series with a title plus category and value data sources. smooth, marker, and explosion fields are optional by chart family.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| title | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"chartTitleSpecTypes"} |  |
| categories | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"chartDataSourceSpecTypes"} |  |
| values | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"chartDataSourceSpecTypes"} |  |
| smooth | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| markerStyle | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| markerSize | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| explosion | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "ChartSeriesSpec",
  "summary" : "One authored chart series with a title plus category and value data sources. smooth, marker, and explosion fields are optional by chart family.",
  "fields" : [ {
    "name" : "title",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "chartTitleSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "categories",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "chartDataSourceSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "values",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "chartDataSourceSpecTypes"
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
