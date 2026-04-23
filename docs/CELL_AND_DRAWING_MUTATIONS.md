---
afad: "3.5"
version: "0.55.0"
domain: CELL_DRAWING_MUTATIONS
updated: "2026-04-22"
route:
  keywords: [gridgrind, cell mutations, drawing mutations, set-cell, set-range, comment, hyperlink, picture, shape, embedded-object, chart, signature-line]
  questions: ["how do I write cells in gridgrind", "how do I set comments or hyperlinks in gridgrind", "how do I add a picture in gridgrind", "how do I author charts in gridgrind", "how do I add a signature line in gridgrind"]
---

# Cell And Drawing Mutation Reference

**Purpose**: Mutation reference for cell content, hyperlinks, comments, pictures, shapes,
embedded objects, charts, signature lines, and drawing-anchor changes.
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[WORKBOOK_AND_LAYOUT_MUTATIONS.md](./WORKBOOK_AND_LAYOUT_MUTATIONS.md),
[STRUCTURED_FEATURE_MUTATIONS.md](./STRUCTURED_FEATURE_MUTATIONS.md), and
[ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)

### SET_CELL

Write a typed value to a single cell using A1 notation. Existing style, hyperlink, and comment
state on the targeted cell is preserved.

```json
{
  "stepId": "set-cell",
  "target": {
    "type": "BY_ADDRESS",
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
    "type": "BY_RANGE",
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
    "type": "BY_RANGE",
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
    "type": "BY_ADDRESS",
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
    "type": "BY_RANGE",
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

### SET_HYPERLINK

Attach or replace a hyperlink on one cell. The cell is created if needed.

```json
{
  "stepId": "set-hyperlink",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Inventory",
    "address": "A1"
  },
  "action": {
    "type": "SET_HYPERLINK",
    "target": {
      "type": "URL",
      "target": "https://example.com/catalog"
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `address` | Yes | A1-notation cell address. |
| `target` | Yes | Typed hyperlink target payload. |

Supported target variants:

- `{"type":"URL","target":"https://example.com/report"}`
- `{"type":"EMAIL","email":"team@example.com"}`
- `{"type":"FILE","path":"shared/report.xlsx"}`
- `{"type":"DOCUMENT","target":"Inventory!B4"}`

`FILE.path` accepts either a plain relative or absolute file path, or a `file:` URI. Read-side
metadata always returns a normalized plain path string in `path`, never a `file:` URI. Relative
`FILE` targets are analyzed against the workbook's persisted path when one exists, so use
absolute paths when you want cwd-independent health checks.

---

### CLEAR_HYPERLINK

Remove any hyperlink attached to one cell. The cell value, style, and comment are left unchanged.
No-op when the cell does not physically exist.

```json
{
  "stepId": "clear-hyperlink",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Inventory",
    "address": "A1"
  },
  "action": {
    "type": "CLEAR_HYPERLINK"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `address` | Yes | A1-notation cell address. |

---

### SET_COMMENT

Attach or replace one comment on one cell. The cell is created if needed. GridGrind always
requires the authoritative plain `text` plus `author`, and can also author ordered rich-text
`runs` and an explicit comment `anchor`.

```json
{
  "stepId": "set-comment",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Inventory",
    "address": "B4"
  },
  "action": {
    "type": "SET_COMMENT",
    "comment": {
      "text": {
        "type": "INLINE",
        "text": "Lead review scheduled."
      },
      "author": "GridGrind",
      "visible": false,
      "runs": [
        {
          "font": {
            "bold": true,
            "fontColorTheme": 4,
            "fontColorTint": -0.2
          },
          "source": {
            "type": "INLINE",
            "text": "Lead"
          }
        },
        {
          "source": {
            "type": "INLINE",
            "text": " "
          }
        },
        {
          "font": {
            "italic": true,
            "fontColorIndexed": 17
          },
          "source": {
            "type": "INLINE",
            "text": "review scheduled."
          }
        }
      ],
      "anchor": {
        "firstColumn": 4,
        "firstRow": 1,
        "lastColumn": 8,
        "lastRow": 7
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `address` | Yes | A1-notation cell address. |
| `comment` | Yes | Comment payload. |

`comment.visible` defaults to `false` when omitted.

`comment` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `text` | Yes | Source-backed plain-text payload. Inline text uses `{ "type": "INLINE", "text": "..." }`; file-backed and standard-input variants are also supported. The resolved text must be nonblank. |
| `author` | Yes | Nonblank comment author. |
| `visible` | No | Whether the comment box is shown by default. Defaults to `false`. |
| `runs` | No | Ordered rich-text run list. When present, the concatenated run text must equal `comment.text` exactly. |
| `anchor` | No | Explicit comment-box bounds using zero-based `firstColumn`, `firstRow`, `lastColumn`, and `lastRow`. |

---

### CLEAR_COMMENT

Remove any comment attached to one cell. The cell value, style, and hyperlink are left unchanged.
No-op when the cell does not physically exist.

```json
{
  "stepId": "clear-comment",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Inventory",
    "address": "B4"
  },
  "action": {
    "type": "CLEAR_COMMENT"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `address` | Yes | A1-notation cell address. |

---

### SET_PICTURE

Create or replace one named picture-backed drawing object on one sheet. Authored drawing mutations
currently accept only `anchor.type = "TWO_CELL"` with zero-based `from` and `to` markers.

```json
{
  "type": "SET_PICTURE",
  "sheetName": "Ops",
  "picture": {
    "name": "OpsPicture",
    "image": {
      "format": "PNG",
      "source": {
        "type": "INLINE_BASE64",
        "base64Data": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="
      }
    },
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 0,
        "rowIndex": 4,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 2,
        "rowIndex": 8,
        "dx": 0,
        "dy": 0
      },
      "behavior": "MOVE_AND_RESIZE"
    },
    "description": {
      "type": "INLINE",
      "text": "Queue preview"
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `picture` | Yes | Authoritative picture payload. |

`picture` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local drawing-object name. Replaces any existing drawing object of the same name. |
| `image` | Yes | Binary picture payload with `format` plus base64 bytes. |
| `anchor` | Yes | Authored drawing anchor. The current public contract supports only `TWO_CELL`. |
| `description` | No | Optional nonblank descriptive text stored on the picture. |

`image.format` accepts `EMF`, `WMF`, `PICT`, `JPEG`, `PNG`, `DIB`, `GIF`, `TIFF`, `EPS`, `BMP`,
or `WPG`.

`anchor` currently uses:

```json
{
  "type": "TWO_CELL",
  "from": {
    "columnIndex": 0,
    "rowIndex": 4,
    "dx": 0,
    "dy": 0
  },
  "to": {
    "columnIndex": 2,
    "rowIndex": 8,
    "dx": 0,
    "dy": 0
  },
  "behavior": "MOVE_AND_RESIZE"
}
```

`from` and `to` markers use zero-based cell indexes plus non-negative raw cell-relative offsets in
`dx` and `dy`. `behavior` accepts `MOVE_AND_RESIZE`, `MOVE_DONT_RESIZE`, or
`DONT_MOVE_AND_RESIZE` and defaults to `MOVE_AND_RESIZE` when omitted.

---

### SET_SHAPE

Create or replace one named simple shape or connector on one sheet.

```json
{
  "type": "SET_SHAPE",
  "sheetName": "Ops",
  "shape": {
    "name": "OpsShape",
    "kind": "SIMPLE_SHAPE",
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 3,
        "rowIndex": 4,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 5,
        "rowIndex": 7,
        "dx": 0,
        "dy": 0
      },
      "behavior": "MOVE_DONT_RESIZE"
    },
    "presetGeometryToken": "roundRect",
    "text": {
      "type": "INLINE",
      "text": "Queue"
    }
  }
}
```

```json
{
  "type": "SET_SHAPE",
  "sheetName": "Ops",
  "shape": {
    "name": "OpsConnector",
    "kind": "CONNECTOR",
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 3,
        "rowIndex": 8,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 6,
        "rowIndex": 9,
        "dx": 0,
        "dy": 0
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `shape` | Yes | Authoritative shape payload. |

`shape` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local drawing-object name. |
| `kind` | Yes | `SIMPLE_SHAPE` or `CONNECTOR`. |
| `anchor` | Yes | Authored drawing anchor. The current public contract supports only `TWO_CELL`. |
| `presetGeometryToken` | No | Optional preset geometry token for `SIMPLE_SHAPE`. Defaults to `rect` when omitted. Not allowed for `CONNECTOR`. |
| `text` | No | Optional nonblank text for `SIMPLE_SHAPE`. Not allowed for `CONNECTOR`. |

Validation failures are non-mutating. If `presetGeometryToken` is unsupported, GridGrind rejects
the request without deleting an existing object of the same name and without leaving a partial new
shape behind.

---

### SET_EMBEDDED_OBJECT

Create or replace one named embedded object with an authoritative package payload and preview
image.

```json
{
  "type": "SET_EMBEDDED_OBJECT",
  "sheetName": "Ops",
  "embeddedObject": {
    "name": "OpsEmbed",
    "label": "Ops payload",
    "fileName": "ops-payload.txt",
    "command": "open",
    "previewImage": {
      "format": "PNG",
      "source": {
        "type": "INLINE_BASE64",
        "base64Data": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="
      }
    },
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 6,
        "rowIndex": 4,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 8,
        "rowIndex": 9,
        "dx": 0,
        "dy": 0
      }
    },
    "payload": {
      "type": "INLINE_BASE64",
      "base64Data": "R3JpZEdyaW5kIGVtYmVkZGVkIHBheWxvYWQ="
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `embeddedObject` | Yes | Authoritative embedded-object payload. |

`embeddedObject` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local drawing-object name. |
| `label` | Yes | Nonblank display label stored on the embedded object. |
| `fileName` | Yes | Nonblank embedded file name surfaced again in payload reads when available. |
| `command` | Yes | Nonblank command string stored on the object metadata. |
| `base64Data` | Yes | Base64-encoded embedded payload bytes. |
| `previewImage` | Yes | Preview picture payload shown by Excel for the object. |
| `anchor` | Yes | Authored drawing anchor. The current public contract supports only `TWO_CELL`. |

---

### SET_CHART

Create or mutate one named supported chart on one sheet. The public chart contract now models one
chart with ordered `plots`, so authored requests can build one supported plot or one multi-plot
combo chart. Supported authored plot families are `AREA`, `AREA_3D`, `BAR`, `BAR_3D`,
`DOUGHNUT`, `LINE`, `LINE_3D`, `PIE`, `PIE_3D`, `RADAR`, `SCATTER`, `SURFACE`, and
`SURFACE_3D`. Series bind to workbook ranges, defined names, or literal values depending on the
plot. Existing unsupported loaded detail is preserved on unrelated edits and rejected for
authoritative mutation.

```json
{
  "type": "SET_CHART",
  "sheetName": "Ops",
  "chart": {
    "name": "OpsChart",
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 4,
        "rowIndex": 0,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 8,
        "rowIndex": 12,
        "dx": 0,
        "dy": 0
      },
      "behavior": "MOVE_AND_RESIZE"
    },
    "title": {
      "type": "TEXT",
      "source": {
        "type": "INLINE",
        "text": "Roadmap"
      }
    },
    "legend": {
      "type": "VISIBLE",
      "position": "TOP_RIGHT"
    },
    "displayBlanksAs": "SPAN",
    "plotOnlyVisibleCells": false,
    "plots": [
      {
        "type": "BAR",
        "varyColors": true,
        "barDirection": "COLUMN",
        "series": [
          {
            "title": {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Plan"
              }
            },
            "categories": {
              "type": "REFERENCE",
              "formula": "ChartCategories"
            },
            "values": {
              "type": "REFERENCE",
              "formula": "Ops!$B$2:$B$4"
            }
          },
          {
            "title": {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Actual"
              }
            },
            "categories": {
              "type": "REFERENCE",
              "formula": "ChartCategories"
            },
            "values": {
              "type": "REFERENCE",
              "formula": "ChartActual"
            }
          }
        ]
      }
    ]
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `chart` | Yes | Authoritative supported-chart payload. |

`chart` common fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local chart name. Reuses the existing chart when the name already exists. |
| `anchor` | Yes | Authored chart-frame anchor. Currently only `TWO_CELL` is supported. |
| `title` | No | `NONE`, `TEXT`, or `FORMULA`. Defaults to `NONE`. |
| `legend` | No | `HIDDEN` or `VISIBLE`. `VISIBLE.position` defaults to `RIGHT` when omitted. |
| `displayBlanksAs` | No | `GAP`, `SPAN`, or `ZERO`. Defaults to `GAP`. |
| `plotOnlyVisibleCells` | No | Whether hidden cells are ignored. Defaults to `true`. |
| `plots` | Yes | Ordered non-empty plot list. Use one plot for a simple chart or multiple plots for a combo chart. |

`plot` common fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `type` | Yes | `AREA`, `AREA_3D`, `BAR`, `BAR_3D`, `DOUGHNUT`, `LINE`, `LINE_3D`, `PIE`, `PIE_3D`, `RADAR`, `SCATTER`, `SURFACE`, or `SURFACE_3D`. |
| `varyColors` | No | Whether Excel varies series colors automatically. Defaults to `false`. |
| `series` | Yes | Ordered non-empty series list. |

Axis-backed plot families (`AREA`, `AREA_3D`, `BAR`, `BAR_3D`, `LINE`, `LINE_3D`, `RADAR`,
`SCATTER`, `SURFACE`, `SURFACE_3D`) may also provide explicit ordered `axes`. When omitted,
GridGrind writes the default axis set for that plot family.

`series` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `title` | No | `NONE`, `TEXT`, or `FORMULA`. Defaults to `NONE`. |
| `categories` | Yes | `REFERENCE`, `STRING_LITERAL`, or `NUMERIC_LITERAL` data source for the category axis. |
| `values` | Yes | `REFERENCE`, `STRING_LITERAL`, or `NUMERIC_LITERAL` data source for the value axis. |
| `smooth` | No | Line and scatter smoothing flag when the target plot family supports it. |
| `markerStyle` | No | Marker shape for line and scatter series. |
| `markerSize` | No | Marker size between `2` and `72`. |
| `explosion` | No | Pie or doughnut slice explosion amount. Must not be negative. |

Variant-specific fields:

| Chart type | Field | Required | Description |
|:-----------|:------|:---------|:------------|
| `AREA`, `AREA_3D`, `LINE`, `LINE_3D` | `grouping` | No | Plot grouping. Defaults to `STANDARD`. |
| `AREA_3D`, `LINE_3D`, `BAR_3D` | `gapDepth` | No | Optional 3D gap depth. |
| `BAR`, `BAR_3D` | `barDirection` | No | `COLUMN` or `BAR`. Defaults to `COLUMN`. |
| `BAR`, `BAR_3D` | `grouping` | No | Bar grouping. Defaults to `CLUSTERED`. |
| `BAR` | `gapWidth`, `overlap` | No | Optional bar spacing overrides. `overlap` must be between `-100` and `100`. |
| `BAR_3D` | `gapWidth`, `shape` | No | Optional bar spacing and 3D bar shape overrides. |
| `DOUGHNUT` | `firstSliceAngle`, `holeSize` | No | `firstSliceAngle` must be between `0` and `360`; `holeSize` must be between `10` and `90`. |
| `PIE` | `firstSliceAngle` | No | Integer between `0` and `360`. |
| `RADAR` | `style` | No | Radar style. Defaults to `STANDARD`. |
| `SCATTER` | `style` | No | Scatter style. Defaults to `LINE_MARKER`. |
| `SURFACE`, `SURFACE_3D` | `wireframe` | No | Whether the surface plot is written as wireframe. Defaults to `false`. |

Reference-backed `categories` and `values` accept either contiguous A1-style references such as
`Ops!$B$2:$B$4` or defined names such as `ChartActual`.
Formula-backed chart titles and series titles must resolve to one cell, either directly or
through a defined name that resolves to one cell.
Validation failures are non-mutating: GridGrind rejects invalid authored chart payloads without
creating a partial chart or partially mutating an existing supported chart.

---

### SET_SIGNATURE_LINE

Create or replace one named signature line on one sheet. Signature lines are VML-backed drawing
objects with authored signer suggestions, comment permissions, optional caption text, and an
optional plain-signature preview image.

```json
{
  "type": "SET_SIGNATURE_LINE",
  "sheetName": "Approvals",
  "signatureLine": {
    "name": "BudgetSignature",
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 1,
        "rowIndex": 1,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 4,
        "rowIndex": 6,
        "dx": 0,
        "dy": 0
      },
      "behavior": "MOVE_AND_RESIZE"
    },
    "allowComments": false,
    "signingInstructions": "Review the budget before signing.",
    "suggestedSigner": "Ada Lovelace",
    "suggestedSigner2": "Finance",
    "suggestedSignerEmail": "ada@example.com",
    "invalidStamp": "invalid",
    "plainSignature": {
      "format": "PNG",
      "source": {
        "type": "INLINE_BASE64",
        "base64Data": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target sheet. |
| `signatureLine` | Yes | Authoritative signature-line payload. |

`signatureLine` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `name` | Yes | Nonblank sheet-local drawing-object name. Replaces any existing signature line of the same name. |
| `anchor` | Yes | Authored drawing anchor. The current public contract supports only `TWO_CELL`. |
| `allowComments` | No | Whether Excel allows comments alongside the signature line. Defaults to `true`. |
| `signingInstructions` | No | Optional nonblank instruction text shown in Excel's signing UI. |
| `suggestedSigner` | No | Optional nonblank signer name. |
| `suggestedSigner2` | No | Optional nonblank second signer line, often used for a title or department. |
| `suggestedSignerEmail` | No | Optional nonblank signer email address. |
| `caption` | No | Optional nonblank caption text limited to at most three lines. |
| `invalidStamp` | No | Optional nonblank stamp text used for invalid-signature rendering. |
| `plainSignature` | No | Optional preview image payload for the plain signature. Reuses the same `format` plus binary `source` shape as `SET_PICTURE.picture.image`. |

At least one signer-suggestion field (`caption`, `suggestedSigner`, `suggestedSigner2`, or
`suggestedSignerEmail`) must be present. Validation failures are non-mutating: invalid
signature-line payloads do not delete an existing signature line or leave a partial VML shape
behind.

---

### SET_DRAWING_OBJECT_ANCHOR

Move one existing named drawing object by replacing its anchor authoritatively. Supported named
targets include pictures, signature lines, simple shapes, connectors, embedded objects, and
supported charts.

```json
{
  "stepId": "set-drawing-object-anchor",
  "target": {
    "type": "BY_NAME",
    "sheetName": "Ops",
    "objectName": "OpsPicture"
  },
  "action": {
    "type": "SET_DRAWING_OBJECT_ANCHOR",
    "anchor": {
      "type": "TWO_CELL",
      "from": {
        "columnIndex": 1,
        "rowIndex": 5,
        "dx": 0,
        "dy": 0
      },
      "to": {
        "columnIndex": 3,
        "rowIndex": 9,
        "dx": 0,
        "dy": 0
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | `BY_NAME` drawing-object selector for the existing sheet-local object. |
| `anchor` | Yes | Replacement authored drawing anchor. The current public contract supports only `TWO_CELL`. |

---

### DELETE_DRAWING_OBJECT

Delete one existing named drawing object from one sheet.

```json
{
  "stepId": "delete-drawing-object",
  "target": {
    "type": "BY_NAME",
    "sheetName": "Ops",
    "objectName": "OpsConnector"
  },
  "action": {
    "type": "DELETE_DRAWING_OBJECT"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | `BY_NAME` drawing-object selector for the existing sheet-local object to delete. Signature-line deletes also remove the associated preview-image relation when one exists. |

---
