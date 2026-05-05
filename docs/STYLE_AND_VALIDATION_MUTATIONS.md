---
afad: "4.0"
version: "0.63.0"
domain: STYLE_VALIDATION_MUTATIONS
updated: "2026-05-01"
route:
  keywords: [gridgrind, style mutations, data validation, conditional formatting, apply-style]
  questions: ["how do i style cells in gridgrind", "how do i set data validation in gridgrind", "how do i manage conditional formatting in gridgrind"]
---

# Style And Validation Mutation Reference

**Purpose**: Detailed mutation reference for cell styles, data validations, and conditional
formatting rules.
**Landing page**: [STRUCTURED_FEATURE_MUTATIONS.md](./STRUCTURED_FEATURE_MUTATIONS.md)
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[STRUCTURED_DATA_MUTATIONS.md](./STRUCTURED_DATA_MUTATIONS.md), and
[ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)

### APPLY_STYLE

Apply a style patch to a rectangular range. Only fields present in `style` are changed;
unspecified style properties are left untouched.

```json
{
  "stepId": "apply-style",
  "target": {
    "type": "RANGE_BY_RANGE",
    "sheetName": "Inventory",
    "range": "A1:C1"
  },
  "action": {
    "type": "APPLY_STYLE",
    "style": {
      "alignment": {
        "wrapText": true,
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "CENTER",
        "textRotation": 15
      },
      "font": {
        "bold": true,
        "fontName": "Aptos",
        "fontHeight": {
          "type": "POINTS",
          "points": 13
        },
        "fontColor": "#FFFFFF"
      },
      "fill": {
        "pattern": "THIN_HORIZONTAL_BANDS",
        "foregroundColor": "#1F4E78",
        "backgroundColor": "#D9E2F3"
      },
      "border": {
        "all": {
          "style": "THIN",
          "color": "#FFFFFF"
        },
        "bottom": {
          "style": "DOUBLE",
          "color": "#D9E2F3"
        }
      },
      "protection": {
        "locked": true
      }
    }
  }
}
```

```json
{
  "stepId": "apply-style",
  "target": {
    "type": "RANGE_BY_RANGE",
    "sheetName": "Inventory",
    "range": "J2:J3"
  },
  "action": {
    "type": "APPLY_STYLE",
    "style": {
      "font": {
        "fontColor": {
          "kind": "THEME",
          "theme": 6,
          "tint": -0.35
        }
      },
      "fill": {
        "kind": "GRADIENT",
        "gradient": {
          "type": "LINEAR",
          "degree": 45.0,
          "stops": [
            {
              "position": 0.0,
              "color": {
                "kind": "RGB",
                "rgb": "#1F497D"
              }
            },
            {
              "position": 1.0,
              "color": {
                "kind": "THEME",
                "theme": 4,
                "tint": 0.45
              }
            }
          ]
        }
      },
      "border": {
        "bottom": {
          "style": "THIN",
          "color": {
            "kind": "INDEXED",
            "indexed": 8
          }
        }
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `range` | Yes | Range to style. |
| `style` | Yes | Partial style patch. `numberFormat` plus nested `alignment`, `font`, `fill`, `border`, and `protection` groups are all optional. |

Style fields:

| Field | Type | Values |
|:------|:-----|:-------|
| `numberFormat` | string | Any Excel number format string, e.g. `"#,##0.00"`, `"yyyy-mm-dd"` |
| `alignment` | object | Optional alignment patch. See below. |
| `font` | object | Optional font patch. See below. |
| `fill` | object | Optional fill patch. See below. |
| `border` | object | Optional border patch. See below. |
| `protection` | object | Optional cell-protection patch. See below. |

Alignment patch:

| Field | Type | Values |
|:------|:-----|:-------|
| `alignment.wrapText` | boolean | `true` / `false` |
| `alignment.horizontalAlignment` | string | `"LEFT"`, `"CENTER"`, `"RIGHT"`, `"GENERAL"` |
| `alignment.verticalAlignment` | string | `"TOP"`, `"CENTER"`, `"BOTTOM"` |
| `alignment.textRotation` | integer | Explicit `0..180` text-rotation degrees |
| `alignment.indentation` | integer | Excel's `0..250` cell-indent scale |

Font patch:

| Field | Type | Values |
|:------|:-----|:-------|
| `font.bold` | boolean | `true` / `false` |
| `font.italic` | boolean | `true` / `false` |
| `font.fontName` | string | Excel font family name, e.g. `"Aptos"` |
| `font.fontHeight` | object | Typed font height. See below. |
| `font.fontColor` | object | Structured `ColorInput` object. See below. |
| `font.underline` | boolean | `true` adds a single underline, `false` removes it |
| `font.strikeout` | boolean | `true` / `false` |

Font-height input:

| Field | Type | Values |
|:------|:-----|:-------|
| `font.fontHeight.type` | string | `"POINTS"` or `"TWIPS"` |
| `font.fontHeight.points` | number | Required when `type` is `"POINTS"`. Must resolve exactly to whole twips, e.g. `11`, `11.5`, `13.25` |
| `font.fontHeight.twips` | integer | Required when `type` is `"TWIPS"`. Positive exact twip value where `20` twips = `1` point |

Structured color input:

| Field | Type | Values |
|:------|:-----|:-------|
| `*.kind` | string | `"RGB"`, `"THEME"`, or `"INDEXED"` |
| `*.rgb` | string | Required when `kind="RGB"`. `#RRGGBB` hex string; lowercase input is normalized to uppercase. |
| `*.theme` | integer | Required when `kind="THEME"`. Theme color index. |
| `*.indexed` | integer | Required when `kind="INDEXED"`. Indexed color slot. |
| `*.tint` | number | Optional tint applied to the chosen base color. |

Color-bearing write fields all reuse the same `ColorInput` object shape:

- `font.fontColor`
- `fill.foregroundColor`
- `fill.backgroundColor`
- `border.*.color`

Alpha channels are not part of the public contract.

Fill patch:

| Field | Type | Values |
|:------|:-----|:-------|
| `fill.kind` | string | `"PATTERN_ONLY"`, `"PATTERN_FOREGROUND"`, `"PATTERN_BACKGROUND"`, `"PATTERN_FOREGROUND_BACKGROUND"`, or `"GRADIENT"` |
| `fill.pattern` | string | Required for patterned variants. `ExcelFillPattern` value such as `"SOLID"`, `"THIN_HORIZONTAL_BANDS"`, `"SQUARES"` |
| `fill.foregroundColor` | object | `ColorInput` object required by foreground-bearing variants |
| `fill.backgroundColor` | object | `ColorInput` object required by background-bearing variants |
| `fill.gradient` | object | `CellGradientFillInput` payload required when `fill.kind="GRADIENT"` |

Fill notes:

- `fill.kind="PATTERN_ONLY"` accepts `fill.pattern` only. `fill.pattern="NONE"` is valid here.
- `fill.kind="PATTERN_FOREGROUND"` requires `fill.pattern` plus `fill.foregroundColor`. Use this
  for solid fills.
- `fill.kind="PATTERN_BACKGROUND"` requires `fill.pattern` plus `fill.backgroundColor`.
- `fill.kind="PATTERN_FOREGROUND_BACKGROUND"` requires `fill.pattern`,
  `fill.foregroundColor`, and `fill.backgroundColor`.
- `fill.kind="PATTERN_BACKGROUND"` and `fill.kind="PATTERN_FOREGROUND_BACKGROUND"` reject
  `fill.pattern="SOLID"`.
- `fill.kind="GRADIENT"` uses `fill.gradient` only and is mutually exclusive with patterned fill
  fields.
- `fill.gradient.type="LINEAR"` accepts `degree` only. `fill.gradient.type="PATH"` accepts
  `left`, `right`, `top`, and `bottom` only. Mixing the two geometry models is rejected as an
  invalid request.

`fill.gradient` fields:

| Field | Type | Values |
|:------|:-----|:-------|
| `type` | string | Gradient type, e.g. `"LINEAR"` or `"PATH"`. Defaults to `"LINEAR"` when omitted. |
| `degree` | number | Optional angle for linear gradients. |
| `left` | number | Optional left offset for path gradients. |
| `right` | number | Optional right offset for path gradients. |
| `top` | number | Optional top offset for path gradients. |
| `bottom` | number | Optional bottom offset for path gradients. |
| `stops` | array | Ordered gradient stops; each stop has `position` in `0.0..1.0` and a structured `color` object using `rgb`, `theme`, `indexed`, and optional `tint`. |

Border patch:

| Field | Type | Values |
|:------|:-----|:-------|
| `all` | object | Optional default for every border side not explicitly set in the same patch |
| `top` | object | Optional side-specific override |
| `right` | object | Optional side-specific override |
| `bottom` | object | Optional side-specific override |
| `left` | object | Optional side-specific override |

Each border-side object can set a visible style, a structured color object, or both:

| Field | Type | Values |
|:------|:-----|:-------|
| `style` | string | `"NONE"`, `"THIN"`, `"MEDIUM"`, `"DASHED"`, `"DOTTED"`, `"THICK"`, `"DOUBLE"`, `"HAIR"`, `"MEDIUM_DASHED"`, `"DASH_DOT"`, `"MEDIUM_DASH_DOT"`, `"DASH_DOT_DOT"`, `"MEDIUM_DASH_DOT_DOT"`, `"SLANTED_DASH_DOT"` |
| `color` | object | Optional `ColorInput` object describing the side color |

Border notes:

- `border` must set at least one of `all`, `top`, `right`, `bottom`, or `left`.
- `border.all` acts as the default style and color for every side not explicitly set in the same patch.
- Explicit side settings override `border.all`.
- A border color requires an effective visible style on that side, either set on the side itself
  or inherited from `border.all`.
- `style="NONE"` does not allow a `color`.
- `color` uses the same structured `ColorInput` object shape as font and fill colors.

Protection patch:

| Field | Type | Values |
|:------|:-----|:-------|
| `protection.locked` | boolean | `true` / `false` |
| `protection.hiddenFormula` | boolean | `true` / `false` |

Protection note:

- These flags matter when sheet protection is enabled.
- Cell snapshot reads (GET_CELLS, GET_WINDOW, GET_SHEET_SCHEMA) report `style.font.fontHeight`
  as a plain object with both `twips` and `points` fields:
  `{"twips": 260, "points": 13}`. This differs from the write-side
  `FontHeightInput` discriminated format.

---

### SET_DATA_VALIDATION

Create or replace one data-validation rule over a rectangular A1-style range. If the new rule
overlaps existing validation coverage, GridGrind normalizes the sheet so the written rule becomes
authoritative on its target cells and any remaining fragments are retained around it.

```json
{
  "stepId": "set-data-validation",
  "target": {
    "type": "RANGE_BY_RANGE",
    "sheetName": "Inventory",
    "range": "B2:B200"
  },
  "action": {
    "type": "SET_DATA_VALIDATION",
    "validation": {
      "rule": {
        "type": "EXPLICIT_LIST",
        "values": [
          "Queued",
          "Done",
          "Blocked"
        ]
      },
      "allowBlank": true,
      "suppressDropDownArrow": false,
      "prompt": {
        "title": {
          "type": "INLINE",
          "text": "Status"
        },
        "text": {
          "type": "INLINE",
          "text": "Pick one workflow state."
        }
      },
      "errorAlert": {
        "style": "STOP",
        "title": {
          "type": "INLINE",
          "text": "Invalid status"
        },
        "text": {
          "type": "INLINE",
          "text": "Use one of the allowed values."
        }
      }
    }
  }
}
```

```json
{
  "stepId": "set-data-validation",
  "target": {
    "type": "RANGE_BY_RANGE",
    "sheetName": "Inventory",
    "range": "C2:C200"
  },
  "action": {
    "type": "SET_DATA_VALIDATION",
    "validation": {
      "rule": {
        "type": "WHOLE_NUMBER",
        "operator": "BETWEEN",
        "formula1": "1",
        "formula2": "5"
      },
      "allowBlank": false,
      "suppressDropDownArrow": false
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `range` | Yes | Rectangular A1-style target range. |
| `validation` | Yes | Supported data-validation definition. |

Supported rule families:

| `validation.rule.type` | Required fields | Notes |
|:-----------------------|:----------------|:------|
| `EXPLICIT_LIST` | `values` | One or more allowed strings. |
| `FORMULA_LIST` | `formula` | Formula-driven list. A leading `=` is accepted and stripped automatically. |
| `WHOLE_NUMBER` | `operator`, `formula1` | `formula2` is required only for `BETWEEN` / `NOT_BETWEEN`. |
| `DECIMAL_NUMBER` | `operator`, `formula1` | `formula2` is required only for `BETWEEN` / `NOT_BETWEEN`. |
| `DATE` | `operator`, `formula1` | `formula2` is required only for `BETWEEN` / `NOT_BETWEEN`. |
| `TIME` | `operator`, `formula1` | `formula2` is required only for `BETWEEN` / `NOT_BETWEEN`. |
| `TEXT_LENGTH` | `operator`, `formula1` | `formula2` is required only for `BETWEEN` / `NOT_BETWEEN`. |
| `CUSTOM_FORMULA` | `formula` | A leading `=` is accepted and stripped automatically. |

Validation fields:

| Field | Description |
|:------|:------------|
| `allowBlank` | Explicitly choose whether empty cells bypass validation. |
| `suppressDropDownArrow` | Explicitly choose whether Excel hides the list dropdown arrow when the rule supports it. |
| `prompt` | Optional prompt box with source-backed `title` and `text` values plus optional `showPromptBox` (defaults to `true`). |
| `errorAlert` | Optional error box with `style`, source-backed `title` / `text`, and optional `showErrorBox` (defaults to `true`). |

Comparison operators use the typed string values:
`BETWEEN`, `NOT_BETWEEN`, `EQUAL`, `NOT_EQUAL`, `GREATER_THAN`, `LESS_THAN`,
`GREATER_OR_EQUAL`, and `LESS_OR_EQUAL`.

---

### CLEAR_DATA_VALIDATIONS

Remove data-validation coverage from one sheet. `ALL` clears every validation structure on the
sheet. `SELECTED` removes only the portions whose stored ranges intersect the supplied A1-style
ranges, preserving any remaining coverage fragments around the cleared area.

```json
{
  "stepId": "clear-data-validations",
  "target": {
    "type": "RANGE_ALL_ON_SHEET",
    "sheetName": "Inventory"
  },
  "action": {
    "type": "CLEAR_DATA_VALIDATIONS"
  }
}
```

```json
{
  "stepId": "clear-data-validations",
  "target": {
    "type": "RANGE_BY_RANGES",
    "sheetName": "Inventory",
    "ranges": [
      "B4",
      "C8:D10"
    ]
  },
  "action": {
    "type": "CLEAR_DATA_VALIDATIONS"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `target` | Yes | `ALL_ON_SHEET` or `BY_RANGES` selector for the target ranges. |

---

### SET_CONDITIONAL_FORMATTING

Creates or replaces one logical conditional-formatting block on the sheet. The block owns one or
more target ranges and one ordered rule list. Authoring supports six rule families:

- `FORMULA_RULE`
- `CELL_VALUE_RULE`
- `COLOR_SCALE_RULE`
- `DATA_BAR_RULE`
- `ICON_SET_RULE`
- `TOP10_RULE`

For `FORMULA_RULE`, `CELL_VALUE_RULE`, and `TOP10_RULE`, the optional `style` payload is
differential, not a whole-cell style patch. Supported differential-style attributes are
`numberFormat`, `bold`, `italic`, `fontHeight`, `fontColor`, `underline`, `strikeout`,
`fillColor`, and per-side differential borders.

```json
{
  "stepId": "set-conditional-formatting",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "SET_CONDITIONAL_FORMATTING",
    "conditionalFormatting": {
      "ranges": [
        "D2:D200"
      ],
      "rules": [
        {
          "type": "CELL_VALUE_RULE",
          "operator": "GREATER_THAN",
          "formula1": "8",
          "stopIfTrue": false,
          "style": {
            "numberFormat": "0.0",
            "fillColor": "#FDE9D9",
            "fontColor": "#9C0006",
            "bold": true
          }
        }
      ]
    }
  }
}
```

```json
{
  "stepId": "set-conditional-formatting",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "SET_CONDITIONAL_FORMATTING",
    "conditionalFormatting": {
      "ranges": [
        "K2:K200"
      ],
      "rules": [
        {
          "type": "COLOR_SCALE_RULE",
          "stopIfTrue": false,
          "thresholds": [
            {
              "type": "MIN"
            },
            {
              "type": "PERCENTILE",
              "value": 50.0
            },
            {
              "type": "MAX"
            }
          ],
          "colors": [
            {
              "rgb": "#AA2211"
            },
            {
              "rgb": "#FFDD55"
            },
            {
              "rgb": "#11CC66"
            }
          ]
        }
      ]
    }
  }
}
```

Rule family summary:

| Rule type | Required fields | Notes |
|:----------|:----------------|:------|
| `FORMULA_RULE` | `formula` | Optional `style`. |
| `CELL_VALUE_RULE` | `operator`, `formula1` | Optional `formula2` for between/not-between operators and optional `style`. |
| `COLOR_SCALE_RULE` | `thresholds`, `colors` | Threshold and color list sizes must match and must contain at least two control points. |
| `DATA_BAR_RULE` | `color`, `widthMin`, `widthMax`, `minThreshold`, `maxThreshold` | `iconOnly` controls whether only the bar glyph is shown. No direction field is exposed. |
| `ICON_SET_RULE` | `iconSet`, `thresholds` | Threshold count must match the chosen icon-set family. |
| `TOP10_RULE` | `rank` | Optional `style`; `percent` and `bottom` control Top/Bottom N vs percent behavior. |

### CLEAR_CONDITIONAL_FORMATTING

Removes conditional-formatting blocks on the sheet that intersect the provided range selector.
`ALL_ON_SHEET` clears the sheet. `BY_RANGES` removes every stored block whose target ranges
intersect any of the selected A1 ranges.

```json
{
  "stepId": "clear-conditional-formatting",
  "target": {
    "type": "RANGE_ALL_ON_SHEET",
    "sheetName": "Inventory"
  },
  "action": {
    "type": "CLEAR_CONDITIONAL_FORMATTING"
  }
}
```

```json
{
  "stepId": "clear-conditional-formatting",
  "target": {
    "type": "RANGE_BY_RANGES",
    "sheetName": "Inventory",
    "ranges": [
      "D2:D200",
      "F2:F20"
    ]
  },
  "action": {
    "type": "CLEAR_CONDITIONAL_FORMATTING"
  }
}
```

---
