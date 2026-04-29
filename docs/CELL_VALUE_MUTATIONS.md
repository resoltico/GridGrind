---
afad: "3.5"
version: "0.61.0"
domain: CELL_VALUE_MUTATIONS
updated: "2026-04-25"
route:
  keywords: [gridgrind, cell mutations, set-cell, set-range, clear-range, array-formula]
  questions: ["how do i write cells in gridgrind", "how do i write ranges in gridgrind", "how do i use array formulas in gridgrind"]
---

# Cell Value Mutation Reference

**Purpose**: Detailed mutation reference for scalar cell writes, rectangular range writes,
array-formula authoring, and value clearing.
**Landing page**: [CELL_AND_DRAWING_MUTATIONS.md](./CELL_AND_DRAWING_MUTATIONS.md)
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[LINK_AND_COMMENT_MUTATIONS.md](./LINK_AND_COMMENT_MUTATIONS.md),
[DRAWING_MUTATIONS.md](./DRAWING_MUTATIONS.md), and
[ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)

### SET_CELL

Write a typed value to a single cell using A1 notation. Existing style, hyperlink, and comment
state on the targeted cell is preserved.

```json
{
  "stepId": "set-cell",
  "target": {
    "type": "CELL_BY_ADDRESS",
    "sheetName": "Inventory",
    "address": "B4"
  },
  "action": {
    "type": "SET_CELL",
    "value": {
      "type": "FORMULA",
      "source": {
        "type": "INLINE",
        "text": "SUM(B2:B3)"
      }
    }
  }
}
```

FORMULA note: a leading `=` is accepted and stripped automatically. `"=SUM(B2:B3)"` and
`"SUM(B2:B3)"` are equivalent. Scalar `FORMULA` values stay scalar only; use
`SET_ARRAY_FORMULA` for contiguous array-formula groups.

DATE / DATE_TIME note: the required Excel number format is applied without discarding any existing
fill, border, font, alignment, or wrap state already present on the cell.

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `address` | Yes | A1-notation cell address. |
| `value` | Yes | Typed cell value. |

---

### SET_RANGE

Write a rectangular grid of typed values. The `rows` matrix must exactly match the dimensions of
`range`. Row count must equal `range` row span; column count of each row must equal `range`
column span. Existing style, hyperlink, and comment state on the targeted cells is preserved.
Any typed value variant accepted by `SET_CELL`, including `RICH_TEXT`, is valid inside `rows`.

```json
{
  "stepId": "set-range",
  "target": {
    "type": "RANGE_BY_RANGE",
    "sheetName": "Inventory",
    "range": "A1:C3"
  },
  "action": {
    "type": "SET_RANGE",
    "rows": [
      [
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Origin"
          }
        },
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Kilos"
          }
        },
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Cost/kg"
          }
        }
      ],
      [
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Ethiopia Yirgacheffe"
          }
        },
        {
          "type": "NUMBER",
          "number": 150
        },
        {
          "type": "NUMBER",
          "number": 8.4
        }
      ],
      [
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Colombia Huila"
          }
        },
        {
          "type": "NUMBER",
          "number": 200
        },
        {
          "type": "NUMBER",
          "number": 7.8
        }
      ]
    ]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `range` | Yes | A1-notation rectangular range (e.g. `A1:C3`). |
| `rows` | Yes | 2-D array of typed cell values. Dimensions must match `range`. |

---

### SET_ARRAY_FORMULA

Author one contiguous single-cell or multi-cell array-formula group. The step target is a
rectangular `RangeSelector.BY_RANGE`; the stored array range is exactly the selector range.
Inline formula text may include or omit a leading `=` or `{=...}` wrapper and GridGrind
normalizes the stored formula text before handing it to POI.

```json
{
  "stepId": "set-array-formula",
  "target": {
    "type": "RANGE_BY_RANGE",
    "sheetName": "Calc",
    "range": "D2:D4"
  },
  "action": {
    "type": "SET_ARRAY_FORMULA",
    "formula": {
      "source": {
        "type": "INLINE",
        "text": "{=B2:B4*C2:C4}"
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Range selector for the stored array-formula group. |
| `range` | Yes | Contiguous A1-notation array-formula range. |
| `formula` | Yes | Array-formula payload. Inline text may include or omit a leading `=` or `{=...}` wrapper. |

---

### CLEAR_ARRAY_FORMULA

Remove the stored array-formula group targeted by any member cell. The step target is a
`CellSelector.BY_ADDRESS`; if the addressed cell is not part of an array-formula group, the step
fails explicitly instead of silently doing nothing.

```json
{
  "stepId": "clear-array-formula",
  "target": {
    "type": "CELL_BY_ADDRESS",
    "sheetName": "Calc",
    "address": "D3"
  },
  "action": {
    "type": "CLEAR_ARRAY_FORMULA"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Cell selector targeting any member of the stored array-formula group. |
| `address` | Yes | One member cell of the stored array-formula group. |

---

### CLEAR_RANGE

Remove all values, styles, hyperlinks, and comments from a rectangular range. Cells that have
never been written are silently skipped; this operation never fails because a cell does not
physically exist.

```json
{
  "stepId": "clear-range",
  "target": {
    "type": "RANGE_BY_RANGE",
    "sheetName": "Inventory",
    "range": "A2:C10"
  },
  "action": {
    "type": "CLEAR_RANGE"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `range` | Yes | Range to clear. |

---
