<!-- Generated from core metadata; do not edit. -->
# chartPlotSpecTypes

## AREA

Authored area-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisSpecType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "AREA",
  "summary" : "Authored area-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "grouping",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "STANDARD", "STACKED", "PERCENT_STACKED" ]
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## AREA_3D

Authored 3D area-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| gapDepth | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisSpecType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "AREA_3D",
  "summary" : "Authored 3D area-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "grouping",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "STANDARD", "STACKED", "PERCENT_STACKED" ]
  }, {
    "name" : "gapDepth",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## BAR

Authored bar-chart plot. Supply explicit barDirection, grouping, and axes. The Java authoring surface exposes one convenience constructor for the standard category/value axis pair.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| barDirection | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| gapWidth | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| overlap | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisSpecType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "BAR",
  "summary" : "Authored bar-chart plot. Supply explicit barDirection, grouping, and axes. The Java authoring surface exposes one convenience constructor for the standard category/value axis pair.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "barDirection",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "COLUMN", "BAR" ]
  }, {
    "name" : "grouping",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "STANDARD", "CLUSTERED", "STACKED", "PERCENT_STACKED" ]
  }, {
    "name" : "gapWidth",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "overlap",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## BAR_3D

Authored 3D bar-chart plot. Supply explicit barDirection, grouping, and axes. The Java authoring surface exposes one convenience constructor for the standard category/value axis pair.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| barDirection | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| gapDepth | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| gapWidth | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| shape | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisSpecType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "BAR_3D",
  "summary" : "Authored 3D bar-chart plot. Supply explicit barDirection, grouping, and axes. The Java authoring surface exposes one convenience constructor for the standard category/value axis pair.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "barDirection",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "COLUMN", "BAR" ]
  }, {
    "name" : "grouping",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "STANDARD", "CLUSTERED", "STACKED", "PERCENT_STACKED" ]
  }, {
    "name" : "gapDepth",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "gapWidth",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "shape",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "BOX", "CONE", "CONE_TO_MAX", "CYLINDER", "PYRAMID", "PYRAMID_TO_MAX" ]
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## DOUGHNUT

Authored doughnut-chart plot. varyColors, firstSliceAngle, and holeSize are optional.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| firstSliceAngle | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| holeSize | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "DOUGHNUT",
  "summary" : "Authored doughnut-chart plot. varyColors, firstSliceAngle, and holeSize are optional.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "firstSliceAngle",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "holeSize",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## LINE

Authored line-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisSpecType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "LINE",
  "summary" : "Authored line-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "grouping",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "STANDARD", "STACKED", "PERCENT_STACKED" ]
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## LINE_3D

Authored 3D line-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| gapDepth | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisSpecType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "LINE_3D",
  "summary" : "Authored 3D line-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "grouping",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "STANDARD", "STACKED", "PERCENT_STACKED" ]
  }, {
    "name" : "gapDepth",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## PIE

Authored pie-chart plot. varyColors and firstSliceAngle are optional.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| firstSliceAngle | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "PIE",
  "summary" : "Authored pie-chart plot. varyColors and firstSliceAngle are optional.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "firstSliceAngle",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## PIE_3D

Authored 3D pie-chart plot.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "PIE_3D",
  "summary" : "Authored 3D pie-chart plot.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## RADAR

Authored radar-chart plot. Supply explicit style and axes, or use the Java convenience constructor for the standard category/value axis pair.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| style | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisSpecType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "RADAR",
  "summary" : "Authored radar-chart plot. Supply explicit style and axes, or use the Java convenience constructor for the standard category/value axis pair.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "style",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "FILLED", "MARKER", "STANDARD" ]
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## SCATTER

Authored scatter-chart plot. Supply explicit style and axes, or use the Java convenience constructor for the standard X/Y axis pair.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| style | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisSpecType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "SCATTER",
  "summary" : "Authored scatter-chart plot. Supply explicit style and axes, or use the Java convenience constructor for the standard X/Y axis pair.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "style",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "LINE", "LINE_MARKER", "MARKER", "NONE", "SMOOTH", "SMOOTH_MARKER" ]
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## SURFACE

Authored surface-chart plot. Supply explicit axes, or use the Java convenience constructor for the standard category/value/series axis set.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| wireframe | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisSpecType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "SURFACE",
  "summary" : "Authored surface-chart plot. Supply explicit axes, or use the Java convenience constructor for the standard category/value/series axis set.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "wireframe",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## SURFACE_3D

Authored 3D surface-chart plot. Supply explicit axes, or use the Java convenience constructor for the standard category/value/series axis set.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| wireframe | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisSpecType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesSpecType"}} |  |

```json
{
  "id" : "SURFACE_3D",
  "summary" : "Authored 3D surface-chart plot. Supply explicit axes, or use the Java convenience constructor for the standard category/value/series axis set.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "wireframe",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
