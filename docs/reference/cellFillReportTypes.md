<!-- Generated from core metadata; do not edit. -->
# cellFillReportTypes

## PATTERN_ONLY

Read one patterned fill with no explicit foreground or background colors.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pattern | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "PATTERN_ONLY",
  "summary" : "Read one patterned fill with no explicit foreground or background colors.",
  "fields" : [ {
    "name" : "pattern",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "NONE", "SOLID", "FINE_DOTS", "ALT_BARS", "SPARSE_DOTS", "THICK_HORIZONTAL_BANDS", "THICK_VERTICAL_BANDS", "THICK_BACKWARD_DIAGONAL", "THICK_FORWARD_DIAGONAL", "BIG_SPOTS", "BRICKS", "THIN_HORIZONTAL_BANDS", "THIN_VERTICAL_BANDS", "THIN_BACKWARD_DIAGONAL", "THIN_FORWARD_DIAGONAL", "SQUARES", "DIAMONDS", "LESS_DOTS", "LEAST_DOTS" ]
  } ],
  "targetSelectors" : [ ]
}
```

## PATTERN_FOREGROUND

Read one patterned fill with one factual foreground color.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pattern | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| foregroundColor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellColorReportTypes"} |  |

```json
{
  "id" : "PATTERN_FOREGROUND",
  "summary" : "Read one patterned fill with one factual foreground color.",
  "fields" : [ {
    "name" : "pattern",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "NONE", "SOLID", "FINE_DOTS", "ALT_BARS", "SPARSE_DOTS", "THICK_HORIZONTAL_BANDS", "THICK_VERTICAL_BANDS", "THICK_BACKWARD_DIAGONAL", "THICK_FORWARD_DIAGONAL", "BIG_SPOTS", "BRICKS", "THIN_HORIZONTAL_BANDS", "THIN_VERTICAL_BANDS", "THIN_BACKWARD_DIAGONAL", "THIN_FORWARD_DIAGONAL", "SQUARES", "DIAMONDS", "LESS_DOTS", "LEAST_DOTS" ]
  }, {
    "name" : "foregroundColor",
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

## PATTERN_BACKGROUND

Read one patterned fill with one factual background color.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pattern | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| backgroundColor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellColorReportTypes"} |  |

```json
{
  "id" : "PATTERN_BACKGROUND",
  "summary" : "Read one patterned fill with one factual background color.",
  "fields" : [ {
    "name" : "pattern",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "NONE", "SOLID", "FINE_DOTS", "ALT_BARS", "SPARSE_DOTS", "THICK_HORIZONTAL_BANDS", "THICK_VERTICAL_BANDS", "THICK_BACKWARD_DIAGONAL", "THICK_FORWARD_DIAGONAL", "BIG_SPOTS", "BRICKS", "THIN_HORIZONTAL_BANDS", "THIN_VERTICAL_BANDS", "THIN_BACKWARD_DIAGONAL", "THIN_FORWARD_DIAGONAL", "SQUARES", "DIAMONDS", "LESS_DOTS", "LEAST_DOTS" ]
  }, {
    "name" : "backgroundColor",
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

## PATTERN_FOREGROUND_BACKGROUND

Read one patterned fill with factual foreground and background colors.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pattern | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| foregroundColor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellColorReportTypes"} |  |
| backgroundColor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellColorReportTypes"} |  |

```json
{
  "id" : "PATTERN_FOREGROUND_BACKGROUND",
  "summary" : "Read one patterned fill with factual foreground and background colors.",
  "fields" : [ {
    "name" : "pattern",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "NONE", "SOLID", "FINE_DOTS", "ALT_BARS", "SPARSE_DOTS", "THICK_HORIZONTAL_BANDS", "THICK_VERTICAL_BANDS", "THICK_BACKWARD_DIAGONAL", "THICK_FORWARD_DIAGONAL", "BIG_SPOTS", "BRICKS", "THIN_HORIZONTAL_BANDS", "THIN_VERTICAL_BANDS", "THIN_BACKWARD_DIAGONAL", "THIN_FORWARD_DIAGONAL", "SQUARES", "DIAMONDS", "LESS_DOTS", "LEAST_DOTS" ]
  }, {
    "name" : "foregroundColor",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellColorReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "backgroundColor",
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

## GRADIENT

Read one factual gradient fill.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| gradient | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellGradientFillReportTypes"} |  |

```json
{
  "id" : "GRADIENT",
  "summary" : "Read one factual gradient fill.",
  "fields" : [ {
    "name" : "gradient",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellGradientFillReportTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
