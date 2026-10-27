<!-- Generated from core metadata; do not edit. -->
# cellFillSpecTypes

## PATTERN_ONLY

Write one patterned fill with no explicit foreground or background colors.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pattern | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "PATTERN_ONLY",
  "summary" : "Write one patterned fill with no explicit foreground or background colors.",
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

Write one patterned fill with one explicit foreground color.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pattern | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| foregroundColor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |

```json
{
  "id" : "PATTERN_FOREGROUND",
  "summary" : "Write one patterned fill with one explicit foreground color.",
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
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## PATTERN_BACKGROUND

Write one patterned fill with one explicit background color.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pattern | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| backgroundColor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |

```json
{
  "id" : "PATTERN_BACKGROUND",
  "summary" : "Write one patterned fill with one explicit background color.",
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
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## PATTERN_FOREGROUND_BACKGROUND

Write one patterned fill with explicit foreground and background colors.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pattern | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| foregroundColor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |
| backgroundColor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |

```json
{
  "id" : "PATTERN_FOREGROUND_BACKGROUND",
  "summary" : "Write one patterned fill with explicit foreground and background colors.",
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
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "backgroundColor",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## GRADIENT

Write one gradient fill patch.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| gradient | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellGradientFillSpecTypes"} |  |

```json
{
  "id" : "GRADIENT",
  "summary" : "Write one gradient fill patch.",
  "fields" : [ {
    "name" : "gradient",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellGradientFillSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
