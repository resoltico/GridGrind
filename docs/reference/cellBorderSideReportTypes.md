<!-- Generated from core metadata; do not edit. -->
# cellBorderSideReportTypes

## NONE

Read one border side with no visible border.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Read one border side with no visible border.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## DEFAULT_COLOR

Read one visible border side using Excel's implicit default color.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| style | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "DEFAULT_COLOR",
  "summary" : "Read one visible border side using Excel's implicit default color.",
  "fields" : [ {
    "name" : "style",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "NONE", "THIN", "MEDIUM", "DASHED", "DOTTED", "THICK", "DOUBLE", "HAIR", "MEDIUM_DASHED", "DASH_DOT", "MEDIUM_DASH_DOT", "DASH_DOT_DOT", "MEDIUM_DASH_DOT_DOT", "SLANTED_DASH_DOT" ]
  } ],
  "targetSelectors" : [ ]
}
```

## COLORED

Read one visible border side with an explicit color reference.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| style | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| color | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellColorReportTypes"} |  |

```json
{
  "id" : "COLORED",
  "summary" : "Read one visible border side with an explicit color reference.",
  "fields" : [ {
    "name" : "style",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "NONE", "THIN", "MEDIUM", "DASHED", "DOTTED", "THICK", "DOUBLE", "HAIR", "MEDIUM_DASHED", "DASH_DOT", "MEDIUM_DASH_DOT", "DASH_DOT_DOT", "MEDIUM_DASH_DOT_DOT", "SLANTED_DASH_DOT" ]
  }, {
    "name" : "color",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellColorReportTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
