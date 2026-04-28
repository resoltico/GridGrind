---
afad: "3.5"
version: "0.60.0"
domain: WORKBOOK_SHEET_MUTATIONS
updated: "2026-04-25"
route:
  keywords: [gridgrind, workbook mutations, sheet mutations, protection, copy-sheet, rename-sheet, custom-xml]
  questions: ["how do i manage sheets in gridgrind", "how do i protect workbooks in gridgrind", "how do i import custom xml mappings in gridgrind"]
---

# Workbook And Sheet Mutation Reference

**Purpose**: Detailed mutation reference for workbook ownership, sheet lifecycle, visibility,
protection, and custom-XML import flows.
**Landing page**: [WORKBOOK_AND_LAYOUT_MUTATIONS.md](./WORKBOOK_AND_LAYOUT_MUTATIONS.md)
**Companion references**: [REQUEST_AND_EXECUTION_REFERENCE.md](./REQUEST_AND_EXECUTION_REFERENCE.md),
[LAYOUT_AND_STRUCTURE_MUTATIONS.md](./LAYOUT_AND_STRUCTURE_MUTATIONS.md), and
[ASSERTION_AND_INSPECTION_REFERENCE.md](./ASSERTION_AND_INSPECTION_REFERENCE.md)

### ENSURE_SHEET

Create a sheet if it does not already exist. Does nothing if the sheet exists.

All mutation operations (`SET_CELL`, `SET_RANGE`, `CLEAR_RANGE`, `APPLY_STYLE`,
`SET_DATA_VALIDATION`, `CLEAR_DATA_VALIDATIONS`, `SET_CONDITIONAL_FORMATTING`,
`CLEAR_CONDITIONAL_FORMATTING`, `SET_HYPERLINK`, `CLEAR_HYPERLINK`, `SET_COMMENT`,
`CLEAR_COMMENT`, `SET_PICTURE`, `SET_CHART`, `SET_SHAPE`, `SET_EMBEDDED_OBJECT`,
`SET_SIGNATURE_LINE`, `SET_DRAWING_OBJECT_ANCHOR`, `DELETE_DRAWING_OBJECT`, `SET_AUTOFILTER`,
`CLEAR_AUTOFILTER`, `APPEND_ROW`, `AUTO_SIZE_COLUMNS`) require the target sheet to already exist.
`SET_TABLE` and `SET_PIVOT_TABLE` also require the target `table.sheetName` or
`pivotTable.sheetName` to exist. Use `ENSURE_SHEET` before the first write to any sheet.

```json
{
  "stepId": "ensure-sheet",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Inventory"
  },
  "action": {
    "type": "ENSURE_SHEET"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |

---

### RENAME_SHEET

Rename an existing sheet. The source sheet must already exist. `newSheetName` must be a valid
Excel sheet name and must not conflict with another sheet name.

```json
{
  "stepId": "rename-sheet",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Archive"
  },
  "action": {
    "type": "RENAME_SHEET",
    "newSheetName": "History"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `newSheetName` | Yes | New sheet name. Must be valid and unique. |

---

### DELETE_SHEET

Delete an existing sheet. A workbook must retain at least one sheet and at least one visible
sheet, so attempting to delete the last remaining sheet or the last visible sheet returns
`INVALID_REQUEST`.

```json
{
  "stepId": "delete-sheet",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Scratch"
  },
  "action": {
    "type": "DELETE_SHEET"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |

---

### MOVE_SHEET

Move an existing sheet to a zero-based workbook position. `targetIndex` is evaluated at the time
the operation runs, after all earlier operations in the same request.

```json
{
  "stepId": "move-sheet",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "History"
  },
  "action": {
    "type": "MOVE_SHEET",
    "targetIndex": 0
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `targetIndex` | Yes | Zero-based destination index. Must be between `0` and `sheetCount - 1`. |

---

### COPY_SHEET

Copy one sheet into a new visible, unselected sheet. The copied sheet is placed either at the end
of workbook order or at an explicit zero-based index. GridGrind preserves supported sheet-local
content such as formulas, validations, conditional formatting, comments, hyperlinks, merged
regions, tables, sheet-scoped names, protection metadata, layout state, and supported drawing
content such as pictures, charts, and embedded objects. GridGrind runs repair passes after POI
sheet cloning so copied drawing relations, copied comments, copied embedded-object worksheet
relation ids, and copied workbook-core structures stay authoritative both in memory and after save
or reopen. Charts backed by defined names stay copy-safe as well: when Apache POI cannot rewrite a
named chart source safely during clone, GridGrind normalizes the copied chart onto explicit
destination-sheet area references while leaving the source-sheet chart formula unchanged.

```json
{
  "type": "COPY_SHEET",
  "newSheetName": "Budget Review",
  "position": {
    "type": "AT_INDEX",
    "targetIndex": 1
  },
  "source": {
    "type": "SHEET_BY_NAME",
    "name": "Budget"
  }
}
```

```json
{
  "type": "COPY_SHEET",
  "newSheetName": "Budget Snapshot",
  "source": {
    "type": "SHEET_BY_NAME",
    "name": "Budget"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `source` | Yes | `BY_NAME` selector for the existing sheet to copy. |
| `newSheetName` | Yes | New unique destination sheet name. |
| `position` | No | Copy position. Omit to append at end. |

`position` variants:

- `{"type":"APPEND_AT_END"}`
- `{"type":"AT_INDEX","targetIndex":1}`

---

### SET_ACTIVE_SHEET

Set the active sheet. Hidden sheets cannot be activated. The active sheet is always selected.

```json
{
  "stepId": "set-active-sheet",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Budget Review"
  },
  "action": {
    "type": "SET_ACTIVE_SHEET"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |

---

### SET_SELECTED_SHEETS

Set the selected visible sheet set. Duplicate or unknown sheet names are rejected. Workbook
summary reads return `selectedSheetNames` in workbook order while preserving the chosen active
sheet as the primary selected tab.

```json
{
  "stepId": "set-selected-sheets",
  "target": {
    "type": "SHEET_BY_NAMES",
    "names": [
      "Budget",
      "Budget Review"
    ]
  },
  "action": {
    "type": "SET_SELECTED_SHEETS"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | `BY_NAMES` selector for one or more existing visible sheets. |

---

### SET_SHEET_VISIBILITY

Set one sheet visibility state. A workbook must retain at least one visible sheet, so hiding the
last visible sheet is rejected.

```json
{
  "stepId": "set-sheet-visibility",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Archive"
  },
  "action": {
    "type": "SET_SHEET_VISIBILITY",
    "visibility": "VERY_HIDDEN"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `visibility` | Yes | `VISIBLE`, `HIDDEN`, or `VERY_HIDDEN`. |

---

### SET_SHEET_PROTECTION

Enable sheet protection with the exact supported lock flags and an optional password. When
`password` is provided, GridGrind writes a password hash into the sheet-protection XML while
preserving the explicit authored lock flags.

```json
{
  "stepId": "set-sheet-protection",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Budget Review"
  },
  "action": {
    "type": "SET_SHEET_PROTECTION",
    "protection": {
      "autoFilterLocked": true,
      "deleteColumnsLocked": true,
      "deleteRowsLocked": true,
      "formatCellsLocked": true,
      "formatColumnsLocked": false,
      "formatRowsLocked": false,
      "insertColumnsLocked": true,
      "insertHyperlinksLocked": true,
      "insertRowsLocked": true,
      "objectsLocked": true,
      "pivotTablesLocked": true,
      "scenariosLocked": true,
      "selectLockedCellsLocked": true,
      "selectUnlockedCellsLocked": false,
      "sortLocked": true
    },
    "password": "Sheet-2026"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |
| `protection` | Yes | Supported lock-flag payload. |
| `password` | No | Optional nonblank sheet-protection password. Omit it to protect the sheet without a stored password hash. |

---

### CLEAR_SHEET_PROTECTION

Disable sheet protection entirely and clear any stored password hash.

```json
{
  "stepId": "clear-sheet-protection",
  "target": {
    "type": "SHEET_BY_NAME",
    "name": "Budget Review"
  },
  "action": {
    "type": "CLEAR_SHEET_PROTECTION"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Selector payload for the target workbook location. |

---

### SET_WORKBOOK_PROTECTION

Enable workbook-level protection with authoritative lock flags and optional workbook or revisions
passwords. Omitted booleans normalize to `false`; omitted passwords clear the corresponding stored
hash.

```json
{
  "type": "SET_WORKBOOK_PROTECTION",
  "protection": {
    "structureLocked": true,
    "windowsLocked": false,
    "revisionsLocked": true,
    "workbookPassword": "Vault-2026",
    "revisionsPassword": "Revisions-2026"
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `protection` | Yes | Workbook protection payload carrying structure/windows/revisions lock flags plus optional passwords. |

`protection` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `structureLocked` | No | Lock workbook structure. Defaults to `false`. |
| `windowsLocked` | No | Lock workbook windows. Defaults to `false`. |
| `revisionsLocked` | No | Lock workbook revisions. Defaults to `false`. |
| `workbookPassword` | No | Optional nonblank workbook-protection password. Omit to clear the workbook password hash. |
| `revisionsPassword` | No | Optional nonblank revisions-protection password. Omit to clear the revisions password hash. |

---

### CLEAR_WORKBOOK_PROTECTION

Disable workbook-level protection entirely and remove any stored workbook or revisions password
hashes.

```json
{
  "type": "CLEAR_WORKBOOK_PROTECTION"
}
```

No additional fields.

---

### IMPORT_CUSTOM_XML_MAPPING

Import XML content into one existing workbook custom-XML mapping. The step target is always
`WorkbookSelector.CURRENT`. GridGrind imports data into an existing mapping definition; it does
not author new workbook map definitions.

```json
{
  "stepId": "import-custom-xml",
  "target": {
    "type": "WORKBOOK_CURRENT"
  },
  "action": {
    "type": "IMPORT_CUSTOM_XML_MAPPING",
    "mapping": {
      "locator": {
        "mapId": 1,
        "name": "CORSO_mapping"
      },
      "xml": {
        "type": "UTF8_FILE",
        "path": "custom-xml-assets/custom-xml-update.xml"
      }
    }
  }
}
```

| Field | Required | Description |
|:------|:---------|:------------|
| `target` | Yes | Always `WorkbookSelector.CURRENT`. |
| `mapping` | Yes | Custom-XML import payload containing the existing mapping locator plus the XML content source. |

`mapping` fields:

| Field | Required | Description |
|:------|:---------|:------------|
| `locator.mapId` | No | Stable workbook map id. Supply `mapId`, `name`, or both. |
| `locator.name` | No | Workbook mapping name. Supply `mapId`, `name`, or both. |
| `xml` | Yes | XML content source. Supports `INLINE`, `UTF8_FILE`, and `STANDARD_INPUT`. |

The locator must resolve to exactly one existing mapping.

---
