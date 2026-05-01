---
afad: "4.0"
version: "0.62.0"
domain: LINK_COMMENT_MUTATIONS
updated: "2026-05-01"
route:
  keywords: [gridgrind, hyperlink mutations, comment mutations, set-hyperlink, clear-hyperlink, set-comment, clear-comment]
  questions: ["how do i set hyperlinks in gridgrind", "how do i clear comments in gridgrind", "how do i manage comments in gridgrind"]
---

# Link And Comment Mutation Reference

**Purpose**: Detailed mutation reference for hyperlink and classic-comment authoring and
clearing.
**Landing page**: [CELL_AND_DRAWING_MUTATIONS.md](./CELL_AND_DRAWING_MUTATIONS.md)
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[CELL_VALUE_MUTATIONS.md](./CELL_VALUE_MUTATIONS.md), [DRAWING_MUTATIONS.md](./DRAWING_MUTATIONS.md),
and [ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)

### SET_HYPERLINK

Attach or replace a hyperlink on one cell. The cell is created if needed.

```json
{
  "stepId": "set-hyperlink",
  "target": {
    "type": "CELL_BY_ADDRESS",
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
    "type": "CELL_BY_ADDRESS",
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
    "type": "CELL_BY_ADDRESS",
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
            "fontColor": {
              "kind": "THEME",
              "theme": 4,
              "tint": -0.2
            }
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
            "fontColor": {
              "kind": "INDEXED",
              "indexed": 17
            }
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

`comment.visible` is explicit on the wire. Use `false` when the comment should stay hidden.

`comment` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `text` | Yes | Source-backed plain-text payload. Inline text uses `{ "type": "INLINE", "text": "..." }`; file-backed and standard-input variants are also supported. The resolved text must be nonblank. |
| `author` | Yes | Nonblank comment author. |
| `visible` | Yes | Whether the comment box is shown by default. Use `false` to keep it hidden. |
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
    "type": "CELL_BY_ADDRESS",
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
