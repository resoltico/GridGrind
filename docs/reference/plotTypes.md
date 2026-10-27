<!-- Generated from core metadata; do not edit. -->
# plotTypes

## AREA

Exact area-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisReportType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "AREA",
  "summary" : "Exact area-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
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
        "group" : "chartAxisReportType"
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## AREA_3D

Exact 3D area-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| gapDepth | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisReportType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "AREA_3D",
  "summary" : "Exact 3D area-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
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
        "group" : "chartAxisReportType"
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## BAR

Exact bar-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| barDirection | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| gapWidth | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| overlap | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisReportType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "BAR",
  "summary" : "Exact bar-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
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
        "group" : "chartAxisReportType"
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## BAR_3D

Exact 3D bar-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| barDirection | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| gapDepth | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| gapWidth | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| shape | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisReportType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "BAR_3D",
  "summary" : "Exact 3D bar-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
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
        "group" : "chartAxisReportType"
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## DOUGHNUT

Exact doughnut-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| firstSliceAngle | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| holeSize | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "DOUGHNUT",
  "summary" : "Exact doughnut-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## LINE

Exact line-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisReportType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "LINE",
  "summary" : "Exact line-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
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
        "group" : "chartAxisReportType"
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## LINE_3D

Exact 3D line-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| grouping | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| gapDepth | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisReportType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "LINE_3D",
  "summary" : "Exact 3D line-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
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
        "group" : "chartAxisReportType"
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## PIE

Exact pie-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| firstSliceAngle | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "PIE",
  "summary" : "Exact pie-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## PIE_3D

Exact 3D pie-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "PIE_3D",
  "summary" : "Exact 3D pie-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "series",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## RADAR

Exact radar-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| style | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisReportType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "RADAR",
  "summary" : "Exact radar-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
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
        "group" : "chartAxisReportType"
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## SCATTER

Exact scatter-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| style | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisReportType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "SCATTER",
  "summary" : "Exact scatter-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
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
        "group" : "chartAxisReportType"
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## SURFACE

Exact surface-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| wireframe | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisReportType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "SURFACE",
  "summary" : "Exact surface-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "wireframe",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisReportType"
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## SURFACE_3D

Exact 3D surface-chart plot report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| varyColors | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| wireframe | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| axes | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartAxisReportType"}} |  |
| series | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartSeriesReportType"}} |  |

```json
{
  "id" : "SURFACE_3D",
  "summary" : "Exact 3D surface-chart plot report.",
  "fields" : [ {
    "name" : "varyColors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "wireframe",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "axes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartAxisReportType"
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
        "group" : "chartSeriesReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## UNSUPPORTED

Exact unsupported chart plot report preserved from the workbook.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| plotTypeToken | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| detail | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "UNSUPPORTED",
  "summary" : "Exact unsupported chart plot report preserved from the workbook.",
  "fields" : [ {
    "name" : "plotTypeToken",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "detail",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
