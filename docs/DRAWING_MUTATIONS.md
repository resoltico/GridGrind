---
afad: "3.5"
version: "0.59.0"
domain: DRAWING_MUTATIONS
updated: "2026-04-25"
route:
  keywords: [gridgrind, drawing mutations, picture, shape, embedded-object, chart, signature-line, anchor]
  questions: ["how do i author drawings in gridgrind", "how do i create charts in gridgrind", "how do i move a drawing object in gridgrind"]
---

# Drawing Mutation Reference

**Purpose**: Detailed mutation reference for pictures, shapes, embedded objects, charts,
signature lines, drawing anchors, and drawing deletion.
**Landing page**: [CELL_AND_DRAWING_MUTATIONS.md](./CELL_AND_DRAWING_MUTATIONS.md)
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[CELL_VALUE_MUTATIONS.md](./CELL_VALUE_MUTATIONS.md),
[LINK_AND_COMMENT_MUTATIONS.md](./LINK_AND_COMMENT_MUTATIONS.md), and
[ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)

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
    "type": "DRAWING_OBJECT_BY_NAME",
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
    "type": "DRAWING_OBJECT_BY_NAME",
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
