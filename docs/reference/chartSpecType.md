<!-- Generated from core metadata; do not edit. -->
# chartSpecType

## ChartSpec

One authored chart with an explicit drawing anchor, chart-level presentation state, and one or more plots. Use ChartSpec.withStandardDisplay(...) from Java authoring when the standard legend and blank-cell display policy is intended.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| anchor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingAnchorSpecTypes"} |  |
| title | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"chartTitleSpecTypes"} |  |
| legend | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"chartLegendSpecTypes"} |  |
| displayBlanksAs | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| plotOnlyVisibleCells | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| plots | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"chartPlotSpecTypes"}} |  |

```json
{
  "id" : "ChartSpec",
  "summary" : "One authored chart with an explicit drawing anchor, chart-level presentation state, and one or more plots. Use ChartSpec.withStandardDisplay(...) from Java authoring when the standard legend and blank-cell display policy is intended.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "anchor",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "drawingAnchorSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "title",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "chartTitleSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "legend",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "chartLegendSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "displayBlanksAs",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "GAP", "SPAN", "ZERO" ]
  }, {
    "name" : "plotOnlyVisibleCells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "plots",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "chartPlotSpecTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
