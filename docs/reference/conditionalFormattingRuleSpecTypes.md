<!-- Generated from core metadata; do not edit. -->
# conditionalFormattingRuleSpecTypes

## FORMULA_RULE

Apply one formula-driven conditional-formatting rule. Supply one differential style.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |
| stopIfTrue | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| style | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"differentialStylePatchType"} |  |

```json
{
  "id" : "FORMULA_RULE",
  "summary" : "Apply one formula-driven conditional-formatting rule. Supply one differential style.",
  "fields" : [ {
    "name" : "formula",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "stopIfTrue",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "style",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "differentialStylePatchType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## CELL_VALUE_RULE

Apply one cell-value comparison conditional-formatting rule. formula2 is used only for BETWEEN and NOT_BETWEEN. Supply one differential style.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| operator | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| formula1 | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| formula2 | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| stopIfTrue | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| style | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"differentialStylePatchType"} |  |

```json
{
  "id" : "CELL_VALUE_RULE",
  "summary" : "Apply one cell-value comparison conditional-formatting rule. formula2 is used only for BETWEEN and NOT_BETWEEN. Supply one differential style.",
  "fields" : [ {
    "name" : "operator",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "BETWEEN", "NOT_BETWEEN", "EQUAL", "NOT_EQUAL", "GREATER_THAN", "LESS_THAN", "GREATER_OR_EQUAL", "LESS_OR_EQUAL" ]
  }, {
    "name" : "formula1",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "formula2",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "stopIfTrue",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "style",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "differentialStylePatchType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## COLOR_SCALE_RULE

Apply one color-scale conditional-formatting rule with ordered thresholds and colors.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| stopIfTrue | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| thresholds | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"conditionalFormattingThresholdSpecTypes"}} |  |
| colors | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"}} |  |

```json
{
  "id" : "COLOR_SCALE_RULE",
  "summary" : "Apply one color-scale conditional-formatting rule with ordered thresholds and colors.",
  "fields" : [ {
    "name" : "stopIfTrue",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "thresholds",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "conditionalFormattingThresholdSpecTypes"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "colors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "colorSpecTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## DATA_BAR_RULE

Apply one data-bar conditional-formatting rule with explicit thresholds and widths.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| stopIfTrue | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| color | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |
| iconOnly | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| widthMin | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"INTEGRAL"},{"type":"NUMBER_RANGE","minimum":0.0,"maximum":100.0}]} |  |
| widthMax | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"INTEGRAL"},{"type":"NUMBER_RANGE","minimum":0.0,"maximum":100.0}]} |  |
| minThreshold | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"conditionalFormattingThresholdSpecTypes"} |  |
| maxThreshold | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"conditionalFormattingThresholdSpecTypes"} |  |

```json
{
  "id" : "DATA_BAR_RULE",
  "summary" : "Apply one data-bar conditional-formatting rule with explicit thresholds and widths.",
  "fields" : [ {
    "name" : "stopIfTrue",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "color",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "iconOnly",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "widthMin",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "INTEGRAL"
      }, {
        "type" : "NUMBER_RANGE",
        "minimum" : 0.0,
        "maximum" : 100.0
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "widthMax",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "INTEGRAL"
      }, {
        "type" : "NUMBER_RANGE",
        "minimum" : 0.0,
        "maximum" : 100.0
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "minThreshold",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "conditionalFormattingThresholdSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "maxThreshold",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "conditionalFormattingThresholdSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## ICON_SET_RULE

Apply one icon-set conditional-formatting rule with authored thresholds.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| stopIfTrue | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| iconSet | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| iconOnly | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| reversed | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| thresholds | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"conditionalFormattingThresholdSpecTypes"}} |  |

```json
{
  "id" : "ICON_SET_RULE",
  "summary" : "Apply one icon-set conditional-formatting rule with authored thresholds.",
  "fields" : [ {
    "name" : "stopIfTrue",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "iconSet",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "GYR_3_ARROW", "GREY_3_ARROWS", "GYR_3_FLAGS", "GYR_3_TRAFFIC_LIGHTS", "GYR_3_TRAFFIC_LIGHTS_BOX", "GYR_3_SHAPES", "GYR_3_SYMBOLS_CIRCLE", "GYR_3_SYMBOLS", "GYR_4_ARROWS", "GREY_4_ARROWS", "RB_4_TRAFFIC_LIGHTS", "RATINGS_4", "GYRB_4_TRAFFIC_LIGHTS", "GYYYR_5_ARROWS", "GREY_5_ARROWS", "RATINGS_5", "QUARTERS_5" ]
  }, {
    "name" : "iconOnly",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "reversed",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "thresholds",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "conditionalFormattingThresholdSpecTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## TOP10_RULE

Apply one top/bottom-N conditional-formatting rule with a differential style.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| stopIfTrue | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| rank | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| percent | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| bottom | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| style | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"differentialStylePatchType"} |  |

```json
{
  "id" : "TOP10_RULE",
  "summary" : "Apply one top/bottom-N conditional-formatting rule with a differential style.",
  "fields" : [ {
    "name" : "stopIfTrue",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "rank",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "percent",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "bottom",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "style",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "differentialStylePatchType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
