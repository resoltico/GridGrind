---
afad: "3.5"
version: "0.51.0"
domain: QUICK_REFERENCE
updated: "2026-04-21"
route:
  keywords: [gridgrind, quick-reference, snippets, json, operations, assertions, reads, introspection, analysis, expect-cell-value, expect-analysis-max-severity, copy-paste, ensure-sheet, rename-sheet, delete-sheet, move-sheet, copy-sheet, set-active-sheet, set-selected-sheets, set-sheet-visibility, set-sheet-protection, clear-sheet-protection, set-workbook-protection, clear-workbook-protection, merge-cells, unmerge-cells, set-column-width, set-row-height, set-sheet-pane, set-sheet-zoom, set-print-layout, clear-print-layout, freeze-panes, split-panes, set-cell, set-range, set-hyperlink, clear-hyperlink, set-comment, clear-comment, set-picture, set-chart, set-pivot-table, set-shape, set-embedded-object, set-signature-line, set-drawing-object-anchor, delete-drawing-object, set-data-validation, clear-data-validations, set-autofilter, clear-autofilter, set-table, delete-table, delete-pivot-table, set-named-range, delete-named-range, apply-style, append-row, clear-range, execution-calculation, evaluate-all, evaluate-targets, clear-caches-only, markrecalculateonopen, get-cells, get-window, get-print-layout, get-package-security, get-workbook-protection, get-data-validations, get-autofilters, get-tables, get-pivot-tables, get-drawing-objects, get-charts, get-drawing-object-payload, get-sheet-schema, analyze-autofilter-health, analyze-table-health, analyze-pivot-table-health, analyze-workbook-findings, source-backed, inline, utf8_file, standard_input, inline_base64, file, ooxml, package-security, encryption, signing, coordinates, rowindex, columnindex, warnings]
  questions: ["gridgrind json snippets", "how do I write a cell in gridgrind", "how do I assert a cell value in gridgrind", "how do I assert workbook health in gridgrind", "gridgrind copy paste examples", "gridgrind copy sheet example", "gridgrind active sheet example", "gridgrind selected sheets example", "gridgrind sheet visibility example", "gridgrind sheet protection example", "gridgrind workbook protection example", "gridgrind package security example", "how do I open an encrypted workbook in gridgrind", "how do I inspect package signatures in gridgrind", "gridgrind hyperlink example", "gridgrind comment example", "gridgrind picture example", "gridgrind signature line example", "gridgrind chart example", "gridgrind pivot table example", "how do I read pivot tables in gridgrind", "how do I lint pivot tables in gridgrind", "gridgrind drawing payload example", "gridgrind table example", "gridgrind autofilter example", "gridgrind named range example", "what do gridgrind reads look like", "which gridgrind fields use A1 versus zero-based indexes", "how do I lint workbook health without saving", "how do I load text from a file in gridgrind", "how do I send binary data to gridgrind"]
---

# Quick Reference

Copy-paste JSON fragments for GridGrind workbook plans and ordered steps. Request snippets show
the canonical `steps[]` envelope. Step snippets show complete `MutationStep`, `AssertionStep`, or
`InspectionStep` objects ready to paste into `steps[]`.

GridGrind supports `.xlsx` workbooks only. Use `.xlsx` paths for `source.path` and
`persistence.path`; `.xls`, `.xlsm`, and `.xlsb` are rejected.

The artifact can emit the current contract directly:

```bash
gridgrind --print-request-template
gridgrind --print-protocol-catalog
gridgrind --print-task-catalog
gridgrind --print-task-plan DASHBOARD
gridgrind --print-goal-plan "monthly sales dashboard with charts"
printf '%s\n' '{"source":{"type":"NEW"},"steps":[]}' | gridgrind --doctor-request
gridgrind --print-example WORKBOOK_HEALTH
```

`--print-protocol-catalog` returns machine-readable field descriptors. Every catalog entry states
which fields are required or optional and, for polymorphic fields, which nested or plain type
group supplies the accepted JSON shape. Mutation, assertion, and inspection entries also publish
`targetSelectors` or `targetSelectorRule` so agents can see the exact allowed target families
before sending a request. `--print-task-catalog` and `--print-task-plan <id>` expose the
contract-owned intent layer, `--print-goal-plan "<goal>"` ranks likely tasks for one freeform
goal, and `--doctor-request` returns a machine-readable lint report for one request without
opening or mutating a workbook. Step kind is inferred from exactly one of `action`, `assertion`,
or `query`; do not send `step.type`. `--print-example <id>` emits one contract-owned generated
example request from the same registry that backs the committed `examples/*.json` fixtures.

---

## Request Skeleton

```json
{
  "protocolVersion": "V1",
  "source": { "type": "NEW" },
  "persistence": { "type": "SAVE_AS", "path": "output.xlsx" },
  "execution": { ... },
  "formulaEnvironment": { ... },
  "steps": []
}
```

Path model:
- `source.path` opens an existing workbook from that path.
- `SAVE_AS.path` writes a new workbook to that path and creates missing parent directories.
- `OVERWRITE` writes back to `source.path`; it does not accept its own `path` field.
- Relative paths in `--request`, `--response`, `source.path`, and `persistence.path` resolve from
  the current working directory.

`formulaEnvironment` is optional. Use it when evaluation needs external workbook bindings,
cached-value fallback for unresolved external references, or template-backed UDFs.

`source.security.password` is optional and applies only to encrypted `EXISTING` workbook sources.
Use it when the `.xlsx` package is password-protected.

`persistence.security` is optional on `SAVE_AS` and `OVERWRITE`. Use it to encrypt or sign the
persisted `.xlsx` package:

```json
{
  "persistence": {
    "type": "SAVE_AS",
    "path": "secured.xlsx",
    "security": {
      "encryption": {
        "password": "GridGrind-2026"
      },
      "signature": {
        "pkcs12Path": "signing-material.p12",
        "keystorePassword": "changeit",
        "keyPassword": "changeit",
        "alias": "gridgrind-signing"
      }
    }
  }
}
```

`execution` is optional. Use it only when the request fits one of GridGrind's low-memory
contracts:

```json
{
  "execution": {
    "mode": {
      "readMode": "EVENT_READ",
      "writeMode": "STREAMING_WRITE"
    },
    "journal": {
      "level": "VERBOSE"
    }
  }
}
```

- `execution.mode.readMode: EVENT_READ` supports only `GET_WORKBOOK_SUMMARY` and
  `GET_SHEET_SUMMARY` (`LIM-019`).
- `execution.mode.writeMode: STREAMING_WRITE` requires `source.type: NEW`, supports only
  `ENSURE_SHEET` and `APPEND_ROW`, and allows only
  `execution.calculation.strategy=DO_NOT_CALCULATE` with optional `markRecalculateOnOpen=true`
  (`LIM-020`).
- Every response carries a structured `journal`, even when the request fails. The top-level phases
  are `validation`, `inputResolution`, `open`, `calculation`, `persistencePhase`, and `close`.
- `execution.journal.level` defaults to `NORMAL`. Use `VERBOSE` for live CLI stderr progress and
  verbose event capture in the JSON response journal.
- `execution.calculation` handles immediate evaluation (`EVALUATE_ALL` or `EVALUATE_TARGETS`),
  cache clearing (`CLEAR_CACHES_ONLY`), and workbook-open recalc flags (`markRecalculateOnOpen`).

## Source-Backed Authored Inputs

Text-bearing mutation fields can load their content inline, from UTF-8 files, or from bound stdin
bytes:

```json
{ "type": "INLINE", "text": "Quarterly note" }
{ "type": "UTF8_FILE", "path": "authored-inputs/note.txt" }
{ "type": "STANDARD_INPUT" }
```

Binary-bearing mutation fields use the matching binary source model:

```json
{ "type": "INLINE_BASE64", "base64Data": "SGVsbG8=" }
{ "type": "FILE", "path": "authored-inputs/payload.bin" }
{ "type": "STANDARD_INPUT" }
```

- Relative `UTF8_FILE` and `FILE` paths resolve from the current working directory.
- The request JSON transport is capped at 16 MiB. Put large authored text and binary content in
  `UTF8_FILE`, `FILE`, or `STANDARD_INPUT` sources instead of inline JSON strings.
- `STANDARD_INPUT` authored values require `--request <path>` on the CLI so stdin can carry the
  authored content instead of the request JSON.
- Source-backed loading failures stop the request in `journal.inputResolution` and surface as
  `INPUT_SOURCE_NOT_FOUND`, `INPUT_SOURCE_UNAVAILABLE`, or `INPUT_SOURCE_IO_ERROR`.
- For a full runnable example, see
  [`examples/source-backed-input-request.json`](../examples/source-backed-input-request.json).

## Formula Environment

```json
{
  "formulaEnvironment": {
    "externalWorkbooks": [
      {
        "workbookName": "rates.xlsx",
        "path": "fixtures/rates.xlsx"
      }
    ],
    "missingWorkbookPolicy": "USE_CACHED_VALUE",
    "udfToolpacks": [
      {
        "name": "math",
        "functions": [
          {
            "name": "DOUBLE",
            "minimumArgumentCount": 1,
            "formulaTemplate": "ARG1*2"
          }
        ]
      }
    ]
  }
}
```

## Coordinate Systems

| Field pattern | Convention |
|:--------------|:-----------|
| `address` | A1 cell address, e.g. `B3` |
| `range` | A1 rectangular range, e.g. `A1:C4` |
| `*RowIndex` | Zero-based row index, e.g. `0 = Excel row 1` |
| `*ColumnIndex` | Zero-based column index, e.g. `0 = Excel column A` |

Validation messages echo both forms inline, for example `firstRowIndex 5 (Excel row 6)` or
`firstColumnIndex 5 (Excel column F)`.

## Checking Workbook Health

Use `ANALYZE_WORKBOOK_FINDINGS` as the primary workbook-health check. Pair it with
`persistence.type=NONE` when you only need findings and do not want a saved workbook:

```json
{
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "NONE"
  },
  "steps": [
    {
      "stepId": "lint",
      "target": {
        "type": "CURRENT"
      },
      "query": {
        "type": "ANALYZE_WORKBOOK_FINDINGS"
      }
    }
  ]
}
```

Successful responses may include a `warnings` array. The current request-phase warning flags
same-request sheet names with spaces referenced in formulas without single quotes. Use
`'Sheet Name'!A1` syntax.

For a compact no-save health batch, see
[`examples/workbook-health-request.json`](../examples/workbook-health-request.json). For a larger
mixed introspection-plus-analysis request, see
[`examples/introspection-analysis-request.json`](../examples/introspection-analysis-request.json).

---

## Source

```json
{
  "type": "NEW"
}
{
  "type": "EXISTING",
  "path": "path/to/file.xlsx"
}
{
  "type": "EXISTING",
  "path": "secured.xlsx",
  "security": {
    "password": "GridGrind-2026"
  }
}
```

Relative `path` values resolve from the current working directory.

## Persistence

```json
{
  "type": "SAVE_AS",
  "path": "path/to/output.xlsx"
}
{
  "type": "OVERWRITE"
}
{
  "type": "SAVE_AS",
  "path": "secured.xlsx",
  "security": {
    "encryption": {
      "password": "GridGrind-2026"
    }
  }
}
```

`SAVE_AS.path` writes a new file. `OVERWRITE` writes back to `source.path` and does not accept a
separate `path` field. `SAVE_AS` creates missing parent directories automatically. Signed
workbooks that are mutated require explicit `persistence.security.signature` re-signing before
they can be persisted. `GET_PACKAGE_SECURITY` is available only on the full-XSSF read path.

---

## ENSURE_SHEET

```json
{
  "stepId": "ensure-sheet",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "ENSURE_SHEET"
  }
}
```

## RENAME_SHEET

```json
{
  "stepId": "rename-sheet",
  "target": {
    "type": "BY_NAME",
    "name": "Archive"
  },
  "action": {
    "type": "RENAME_SHEET",
    "newSheetName": "History"
  }
}
```

## DELETE_SHEET

```json
{
  "stepId": "delete-sheet",
  "target": {
    "type": "BY_NAME",
    "name": "Scratch"
  },
  "action": {
    "type": "DELETE_SHEET"
  }
}
```

`DELETE_SHEET` rejects deleting the last remaining sheet or the last visible sheet in a workbook.

## MOVE_SHEET

```json
{
  "stepId": "move-sheet",
  "target": {
    "type": "BY_NAME",
    "name": "History"
  },
  "action": {
    "type": "MOVE_SHEET",
    "targetIndex": 0
  }
}
```

## COPY_SHEET

```json
{
  "type": "COPY_SHEET",
  "newSheetName": "Budget Review",
  "position": {
    "type": "AT_INDEX",
    "targetIndex": 1
  },
  "source": {
    "type": "BY_NAME",
    "name": "Budget"
  }
}
```

Sheet copy preserves non-drawing workbook-core structures such as tables, validations,
conditional formatting, comments, hyperlinks, local names, protection metadata, and print layout.
Drawing-family content such as pictures and charts remains outside the current copy contract.

## SET_ACTIVE_SHEET

```json
{
  "stepId": "set-active-sheet",
  "target": {
    "type": "BY_NAME",
    "name": "Budget Review"
  },
  "action": {
    "type": "SET_ACTIVE_SHEET"
  }
}
```

## SET_SELECTED_SHEETS

```json
{
  "stepId": "set-selected-sheets",
  "target": {
    "type": "BY_NAMES",
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

## SET_SHEET_VISIBILITY

```json
{
  "stepId": "set-sheet-visibility",
  "target": {
    "type": "BY_NAME",
    "name": "Archive"
  },
  "action": {
    "type": "SET_SHEET_VISIBILITY",
    "visibility": "VERY_HIDDEN"
  }
}
```

## SET_SHEET_PROTECTION

```json
{
  "stepId": "set-sheet-protection",
  "target": {
    "type": "BY_NAME",
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

## CLEAR_SHEET_PROTECTION

```json
{
  "stepId": "clear-sheet-protection",
  "target": {
    "type": "BY_NAME",
    "name": "Budget Review"
  },
  "action": {
    "type": "CLEAR_SHEET_PROTECTION"
  }
}
```

## SET_WORKBOOK_PROTECTION

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

## CLEAR_WORKBOOK_PROTECTION

```json
{
  "type": "CLEAR_WORKBOOK_PROTECTION"
}
```

## IMPORT_CUSTOM_XML_MAPPING

```json
{
  "stepId": "import-custom-xml",
  "target": {
    "type": "CURRENT"
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
        "path": "examples/custom-xml-assets/custom-xml-update.xml"
      }
    }
  }
}
```

Workbook custom-XML mappings must already exist in the source workbook. GridGrind imports XML
content into existing mappings; it does not author new map definitions.

## MERGE_CELLS

```json
{
  "stepId": "merge-cells",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Sheet1",
    "range": "A1:C1"
  },
  "action": {
    "type": "MERGE_CELLS"
  }
}
```

## UNMERGE_CELLS

```json
{
  "stepId": "unmerge-cells",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Sheet1",
    "range": "A1:C1"
  },
  "action": {
    "type": "UNMERGE_CELLS"
  }
}
```

## SET_COLUMN_WIDTH

```json
{
  "type": "SET_COLUMN_WIDTH",
  "sheetName": "Sheet1",
  "firstColumnIndex": 0,
  "lastColumnIndex": 2,
  "widthCharacters": 16.0
}
```

## SET_ROW_HEIGHT

```json
{
  "type": "SET_ROW_HEIGHT",
  "sheetName": "Sheet1",
  "firstRowIndex": 0,
  "lastRowIndex": 3,
  "heightPoints": 28.5
}
```

## INSERT_ROWS

```json
{
  "type": "INSERT_ROWS",
  "sheetName": "Sheet1",
  "rowIndex": 2,
  "rowCount": 3
}
```

Append-edge inserts on sparse sheets do not create a physical tail row until content or row
metadata exists there.

## DELETE_ROWS

```json
{
  "type": "DELETE_ROWS",
  "sheetName": "Sheet1",
  "rows": {
    "type": "BAND",
    "firstRowIndex": 4,
    "lastRowIndex": 6
  }
}
```

Rejects deletes that would truncate a range-backed named range (`LIM-018`).

## SHIFT_ROWS

```json
{
  "type": "SHIFT_ROWS",
  "sheetName": "Sheet1",
  "rows": {
    "type": "BAND",
    "firstRowIndex": 1,
    "lastRowIndex": 3
  },
  "delta": 2
}
```

Rejects shifts that would partially move or overwrite a range-backed named range outside the moved
band (`LIM-018`).

## INSERT_COLUMNS

```json
{
  "type": "INSERT_COLUMNS",
  "sheetName": "Sheet1",
  "columnIndex": 1,
  "columnCount": 2
}
```

Append-edge inserts on sparse sheets do not create a physical tail column until cells or explicit
column metadata exist there.

## DELETE_COLUMNS

```json
{
  "type": "DELETE_COLUMNS",
  "sheetName": "Sheet1",
  "columns": {
    "type": "BAND",
    "firstColumnIndex": 3,
    "lastColumnIndex": 4
  }
}
```

Rejects deletes that would truncate a range-backed named range (`LIM-018`) and also rejects
formula-bearing workbooks (`LIM-017`).

## SHIFT_COLUMNS

```json
{
  "type": "SHIFT_COLUMNS",
  "sheetName": "Sheet1",
  "columns": {
    "type": "BAND",
    "firstColumnIndex": 0,
    "lastColumnIndex": 1
  },
  "delta": -1
}
```

Rejects shifts that would partially move or overwrite a range-backed named range outside the moved
band (`LIM-018`) and also rejects formula-bearing workbooks (`LIM-017`).

## SET_ROW_VISIBILITY

```json
{
  "type": "SET_ROW_VISIBILITY",
  "sheetName": "Sheet1",
  "rows": {
    "type": "BAND",
    "firstRowIndex": 5,
    "lastRowIndex": 7
  },
  "hidden": true
}
```

## SET_COLUMN_VISIBILITY

```json
{
  "type": "SET_COLUMN_VISIBILITY",
  "sheetName": "Sheet1",
  "columns": {
    "type": "BAND",
    "firstColumnIndex": 2,
    "lastColumnIndex": 3
  },
  "hidden": false
}
```

## GROUP_ROWS

```json
{
  "type": "GROUP_ROWS",
  "sheetName": "Sheet1",
  "rows": {
    "type": "BAND",
    "firstRowIndex": 8,
    "lastRowIndex": 10
  },
  "collapsed": true
}
```

## UNGROUP_ROWS

```json
{
  "type": "UNGROUP_ROWS",
  "sheetName": "Sheet1",
  "rows": {
    "type": "BAND",
    "firstRowIndex": 8,
    "lastRowIndex": 10
  }
}
```

## GROUP_COLUMNS

```json
{
  "type": "GROUP_COLUMNS",
  "sheetName": "Sheet1",
  "columns": {
    "type": "BAND",
    "firstColumnIndex": 4,
    "lastColumnIndex": 6
  },
  "collapsed": true
}
```

## UNGROUP_COLUMNS

```json
{
  "type": "UNGROUP_COLUMNS",
  "sheetName": "Sheet1",
  "columns": {
    "type": "BAND",
    "firstColumnIndex": 4,
    "lastColumnIndex": 6
  }
}
```

## SET_SHEET_PANE

```json
{
  "stepId": "set-sheet-pane",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "SET_SHEET_PANE",
    "pane": {
      "type": "FROZEN",
      "splitColumn": 1,
      "splitRow": 1,
      "leftmostColumn": 1,
      "topRow": 1
    }
  }
}
```

```json
{
  "stepId": "set-sheet-pane",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "SET_SHEET_PANE",
    "pane": {
      "type": "SPLIT",
      "xSplitPosition": 2400,
      "ySplitPosition": 1800,
      "leftmostColumn": 2,
      "topRow": 3,
      "activePane": "LOWER_RIGHT"
    }
  }
}
```

## SET_SHEET_ZOOM

```json
{
  "stepId": "set-sheet-zoom",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "SET_SHEET_ZOOM",
    "zoomPercent": 125
  }
}
```

## SET_SHEET_PRESENTATION

```json
{
  "stepId": "set-sheet-presentation",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "SET_SHEET_PRESENTATION",
    "presentation": {
      "display": {
        "displayGridlines": false,
        "displayZeros": false,
        "displayRowColHeadings": true,
        "displayFormulas": false,
        "rightToLeft": false
      },
      "tabColor": {
        "rgb": "#0B6E4F"
      },
      "outlineSummary": {
        "rowSumsBelow": false,
        "rowSumsRight": true
      },
      "sheetDefaults": {
        "defaultColumnWidth": 14,
        "defaultRowHeightPoints": 19.0
      },
      "ignoredErrors": [
        {
          "range": "B4:B18",
          "errorTypes": [
            "NUMBER_STORED_AS_TEXT"
          ]
        }
      ]
    }
  }
}
```

`sheetDefaults.defaultColumnWidth` is an integer cell-count width, so fractional JSON numbers are
rejected instead of being rounded or truncated.

## SET_PRINT_LAYOUT

```json
{
  "stepId": "set-print-layout",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "SET_PRINT_LAYOUT",
    "printLayout": {
      "printArea": {
        "type": "RANGE",
        "range": "A1:F40"
      },
      "orientation": "LANDSCAPE",
      "scaling": {
        "type": "FIT",
        "widthPages": 1,
        "heightPages": 0
      },
      "repeatingRows": {
        "type": "BAND",
        "firstRowIndex": 0,
        "lastRowIndex": 1
      },
      "repeatingColumns": {
        "type": "BAND",
        "firstColumnIndex": 0,
        "lastColumnIndex": 0
      },
      "header": {
        "left": {
          "type": "INLINE",
          "text": "Inventory"
        },
        "center": {
          "type": "INLINE",
          "text": "Q2"
        },
        "right": {
          "type": "INLINE",
          "text": "Internal"
        }
      },
      "footer": {
        "left": {
          "type": "INLINE",
          "text": ""
        },
        "center": {
          "type": "INLINE",
          "text": "Prepared by GridGrind"
        },
        "right": {
          "type": "INLINE",
          "text": "Page 1"
        }
      },
      "setup": {
        "margins": {
          "left": 0.35,
          "right": 0.55,
          "top": 0.6,
          "bottom": 0.45,
          "header": 0.3,
          "footer": 0.3
        },
        "horizontallyCentered": true,
        "verticallyCentered": true,
        "printGridlines": true,
        "paperSize": 8,
        "draft": true,
        "blackAndWhite": true,
        "copies": 2,
        "useFirstPageNumber": true,
        "firstPageNumber": 4,
        "rowBreaks": [
          6
        ],
        "columnBreaks": [
          3
        ]
      }
    }
  }
}
```

## CLEAR_PRINT_LAYOUT

```json
{
  "stepId": "clear-print-layout",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "CLEAR_PRINT_LAYOUT"
  }
}
```

## SET_CELL

```json
{
  "stepId": "set-cell",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "A1"
  },
  "action": {
    "type": "SET_CELL",
    "value": {
      "type": "TEXT",
      "source": {
        "type": "INLINE",
        "text": "Hello"
      }
    }
  }
}
{
  "stepId": "set-cell",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "B1"
  },
  "action": {
    "type": "SET_CELL",
    "value": {
      "type": "NUMBER",
      "number": 42.0
    }
  }
}
{
  "stepId": "set-cell",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "C1"
  },
  "action": {
    "type": "SET_CELL",
    "value": {
      "type": "BOOLEAN",
      "bool": true
    }
  }
}
{
  "stepId": "set-cell",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "D1"
  },
  "action": {
    "type": "SET_CELL",
    "value": {
      "type": "FORMULA",
      "source": {
        "type": "INLINE",
        "text": "SUM(B1:B10)"
      }
    }
  }
}
{
  "stepId": "set-cell",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "E1"
  },
  "action": {
    "type": "SET_CELL",
    "value": {
      "type": "DATE",
      "date": "2026-03-25"
    }
  }
}
{
  "stepId": "set-cell",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "F1"
  },
  "action": {
    "type": "SET_CELL",
    "value": {
      "type": "DATE_TIME",
      "dateTime": "2026-03-25T10:15:30"
    }
  }
}
{
  "stepId": "set-cell",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "G1"
  },
  "action": {
    "type": "SET_CELL",
    "value": {
      "type": "BLANK"
    }
  }
}
{
  "stepId": "set-cell",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "H1"
  },
  "action": {
    "type": "SET_CELL",
    "value": {
      "type": "RICH_TEXT",
      "runs": [
        {
          "text": "Q2 "
        },
        {
          "text": "Budget",
          "font": {
            "bold": true,
            "fontColor": "#C00000"
          }
        }
      ]
    }
  }
}
```

`CellInput.TEXT` requires non-empty resolved text. Use `CellInput.BLANK` when you want an empty
cell instead of a string payload.

A leading `=` in `FORMULA` values is accepted and stripped automatically. `"=SUM(B1:B10)"` and `"SUM(B1:B10)"` are equivalent.
Existing style, hyperlink, and comment state on the targeted cell is preserved. `DATE` and
`DATE_TIME` writes keep existing presentation state and only layer the required number format on
top.
`FORMULA` payloads are scalar only. Use `SET_ARRAY_FORMULA` for contiguous single-cell or
multi-cell array-formula groups. Scalar `FORMULA` values that include array-formula braces such
as `{=SUM(A1:A2*B1:B2)}` are rejected as `INVALID_FORMULA`. `LAMBDA` and `LET` are currently
rejected as `INVALID_FORMULA` because Apache POI cannot parse them. Other newer Excel constructs
may fail the same way. Loaded formulas that POI parses but cannot evaluate surface as
`UNSUPPORTED_FORMULA`.

## SET_RANGE

```json
{
  "stepId": "set-range",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Sheet1",
    "range": "A1:C2"
  },
  "action": {
    "type": "SET_RANGE",
    "rows": [
      [
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "A"
          }
        },
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "B"
          }
        },
        {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "C"
          }
        }
      ],
      [
        {
          "type": "NUMBER",
          "number": 1
        },
        {
          "type": "NUMBER",
          "number": 2
        },
        {
          "type": "NUMBER",
          "number": 3
        }
      ]
    ]
  }
}
```

## SET_ARRAY_FORMULA

```json
{
  "stepId": "set-array-formula",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Sheet1",
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

`SET_ARRAY_FORMULA` authors one contiguous array-formula group. Inline formula text may include
or omit a leading `=` or `{=...}` wrapper; GridGrind normalizes the stored formula text.

## CLEAR_ARRAY_FORMULA

```json
{
  "stepId": "clear-array-formula",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "D3"
  },
  "action": {
    "type": "CLEAR_ARRAY_FORMULA"
  }
}
```

`CLEAR_ARRAY_FORMULA` may target any member cell of the stored group and removes the whole
array-formula group.

## CLEAR_RANGE

```json
{
  "stepId": "clear-range",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Sheet1",
    "range": "A1:Z100"
  },
  "action": {
    "type": "CLEAR_RANGE"
  }
}
```

`CLEAR_RANGE` removes value, style, hyperlink, and comment state from every cell in the range.
Cells that have never been written are silently skipped.

## SET_HYPERLINK

```json
{
  "stepId": "set-hyperlink",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "A1"
  },
  "action": {
    "type": "SET_HYPERLINK",
    "target": {
      "type": "URL",
      "target": "https://example.com/report"
    }
  }
}
{
  "stepId": "set-hyperlink",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "A2"
  },
  "action": {
    "type": "SET_HYPERLINK",
    "target": {
      "type": "EMAIL",
      "email": "team@example.com"
    }
  }
}
{
  "stepId": "set-hyperlink",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "A3"
  },
  "action": {
    "type": "SET_HYPERLINK",
    "target": {
      "type": "FILE",
      "path": "reports/monthly.xlsx"
    }
  }
}
{
  "stepId": "set-hyperlink",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "A4"
  },
  "action": {
    "type": "SET_HYPERLINK",
    "target": {
      "type": "DOCUMENT",
      "target": "Summary!B4"
    }
  }
}
```

`FILE.path` accepts either a plain path or a `file:` URI. Reads return a normalized plain path
string in `path`. Relative `FILE` targets are analyzed against the workbook's persisted path when
one exists, so use absolute paths when you want cwd-independent health checks.

## CLEAR_HYPERLINK

```json
{
  "stepId": "clear-hyperlink",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "A1"
  },
  "action": {
    "type": "CLEAR_HYPERLINK"
  }
}
```

## SET_COMMENT

```json
{
  "stepId": "set-comment",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "B2"
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

If `visible` is omitted, it defaults to `false`. When `runs` are present, their text must
concatenate exactly to `comment.text`.

## CLEAR_COMMENT

```json
{
  "stepId": "clear-comment",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Sheet1",
    "address": "B2"
  },
  "action": {
    "type": "CLEAR_COMMENT"
  }
}
```

## SET_PICTURE

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

Authored drawing mutations currently accept only `TWO_CELL` anchors with zero-based `from` and
`to` markers. `behavior` defaults to `MOVE_AND_RESIZE` when omitted.

`image.format` accepts `EMF`, `WMF`, `PICT`, `JPEG`, `PNG`, `DIB`, `GIF`, `TIFF`, `EPS`, `BMP`,
or `WPG`.

## SET_SHAPE

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
    "text": "Queue"
  }
}
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

`kind` is `SIMPLE_SHAPE` or `CONNECTOR`. `presetGeometryToken` and `text` are only for
`SIMPLE_SHAPE`, and `presetGeometryToken` defaults to `rect` when omitted.
Failed validation is non-mutating: unsupported preset geometry does not delete an existing object
or leave a partial new shape behind.

## SET_EMBEDDED_OBJECT

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

`previewImage` reuses the same `format` plus `base64Data` shape as `SET_PICTURE.picture.image`.

## SET_SIGNATURE_LINE

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
      }
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

`allowComments` defaults to `true`. `caption` is optional but limited to three lines, and at least
one signer-suggestion field (`caption`, `suggestedSigner`, `suggestedSigner2`, or
`suggestedSignerEmail`) must be present.

## SET_CHART

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

Formula-backed chart titles and series titles must resolve to one cell, either directly or through
one defined name that resolves to one cell. Reference-backed `categories` and `values` may still
target one contiguous range or one defined name that resolves to one contiguous range. Failed
validation is non-mutating: invalid chart payloads do not create partial charts or half-mutate an
existing supported chart.

Supported authored plot families are `AREA`, `AREA_3D`, `BAR`, `BAR_3D`, `DOUGHNUT`, `LINE`,
`LINE_3D`, `PIE`, `PIE_3D`, `RADAR`, `SCATTER`, `SURFACE`, and `SURFACE_3D`. Authored chart
anchors currently accept only `TWO_CELL`, and one chart may carry multiple ordered `plots` to
build a combo chart.

## SET_DRAWING_OBJECT_ANCHOR

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

## DELETE_DRAWING_OBJECT

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

## APPLY_STYLE

```json
{
  "stepId": "apply-style",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Sheet1",
    "range": "A1:C1"
  },
  "action": {
    "type": "APPLY_STYLE",
    "style": {
      "numberFormat": "#,##0.00",
      "alignment": {
        "wrapText": true,
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "CENTER",
        "textRotation": 15,
        "indentation": 1
      },
      "font": {
        "bold": true,
        "italic": false,
        "fontName": "Aptos",
        "fontHeight": {
          "type": "POINTS",
          "points": 13
        },
        "fontColor": "#1F4E78",
        "underline": true,
        "strikeout": false
      },
      "fill": {
        "pattern": "THIN_HORIZONTAL_BANDS",
        "foregroundColor": "#FFF2CC",
        "backgroundColor": "#FDE9D9"
      },
      "border": {
        "all": {
          "style": "THIN",
          "color": "#D6B656"
        },
        "right": {
          "style": "DOUBLE",
          "color": "#C55A11"
        }
      },
      "protection": {
        "locked": true,
        "hiddenFormula": false
      }
    }
  }
}
```

```json
{
  "stepId": "apply-style",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Sheet1",
    "range": "J2:J3"
  },
  "action": {
    "type": "APPLY_STYLE",
    "style": {
      "font": {
        "fontColorTheme": 6,
        "fontColorTint": -0.35
      },
      "fill": {
        "gradient": {
          "type": "LINEAR",
          "degree": 45.0,
          "stops": [
            {
              "position": 0.0,
              "color": {
                "rgb": "#1F497D"
              }
            },
            {
              "position": 1.0,
              "color": {
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
          "colorIndexed": 8
        }
      }
    }
  }
}
```

`style` groups are `numberFormat`, `alignment`, `font`, `fill`, `border`, and `protection`.
`alignment.horizontalAlignment` values: `"LEFT"` `"CENTER"` `"RIGHT"` `"GENERAL"`
`alignment.verticalAlignment` values: `"TOP"` `"CENTER"` `"BOTTOM"`
`alignment.textRotation` uses explicit `0..180` degrees. `alignment.indentation` uses Excel's `0..250` cell-indent scale.
`font.fontHeight` accepts either `{ "type": "POINTS", "points": 11.5 }` or `{ "type": "TWIPS", "twips": 230 }`.
Color-bearing write fields accept RGB (`fontColor`, `fill.foregroundColor`, `fill.backgroundColor`, `border.*.color`), theme (`fontColorTheme`, `fill.foregroundColorTheme`, `fill.backgroundColorTheme`, `border.*.colorTheme`), or indexed (`fontColorIndexed`, `fill.foregroundColorIndexed`, `fill.backgroundColorIndexed`, `border.*.colorIndexed`) bases plus optional tint companions.
`fill.backgroundColor` and its theme/indexed variants are for patterned fills only; `SOLID` fills use foreground fields only.
`fill.gradient` authors gradient fills and is mutually exclusive with patterned fill fields.
`fill.gradient.type="LINEAR"` accepts `degree` only. `fill.gradient.type="PATH"` accepts `left`, `right`, `top`, and `bottom` only. Mixing the two geometry models is rejected.
`border.all` sets the default side style or color; explicit sides override it.
`border.*.color` requires a visible style on that side, either directly or via `border.all`.

## SET_DATA_VALIDATION

```json
{
  "stepId": "set-data-validation",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Sheet1",
    "range": "B2:B20"
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
{
  "stepId": "set-data-validation",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Sheet1",
    "range": "C2:C20"
  },
  "action": {
    "type": "SET_DATA_VALIDATION",
    "validation": {
      "rule": {
        "type": "WHOLE_NUMBER",
        "operator": "BETWEEN",
        "formula1": "1",
        "formula2": "5"
      }
    }
  }
}
```

Overlapping existing rules are normalized so the written rule becomes authoritative on the target
range.

## CLEAR_DATA_VALIDATIONS

```json
{
  "stepId": "clear-data-validations",
  "target": {
    "type": "ALL_ON_SHEET",
    "sheetName": "Sheet1"
  },
  "action": {
    "type": "CLEAR_DATA_VALIDATIONS"
  }
}
{
  "stepId": "clear-data-validations",
  "target": {
    "type": "BY_RANGES",
    "sheetName": "Sheet1",
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

## SET_CONDITIONAL_FORMATTING

```json
{
  "stepId": "set-conditional-formatting",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "SET_CONDITIONAL_FORMATTING",
    "conditionalFormatting": {
      "ranges": [
        "D2:D20"
      ],
      "rules": [
        {
          "type": "CELL_VALUE_RULE",
          "operator": "GREATER_THAN",
          "formula1": "8",
          "stopIfTrue": false,
          "style": {
            "fillColor": "#FDE9D9",
            "fontColor": "#9C0006",
            "bold": true
          }
        }
      ]
    }
  }
}
{
  "stepId": "set-conditional-formatting",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "SET_CONDITIONAL_FORMATTING",
    "conditionalFormatting": {
      "ranges": [
        "K2:K20"
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

Write support covers `FORMULA_RULE`, `CELL_VALUE_RULE`, `COLOR_SCALE_RULE`, `DATA_BAR_RULE`,
`ICON_SET_RULE`, and `TOP10_RULE`.

## CLEAR_CONDITIONAL_FORMATTING

```json
{
  "stepId": "clear-conditional-formatting",
  "target": {
    "type": "ALL_ON_SHEET",
    "sheetName": "Sheet1"
  },
  "action": {
    "type": "CLEAR_CONDITIONAL_FORMATTING"
  }
}
{
  "stepId": "clear-conditional-formatting",
  "target": {
    "type": "BY_RANGES",
    "sheetName": "Sheet1",
    "ranges": [
      "D2:D20",
      "F2:F20"
    ]
  },
  "action": {
    "type": "CLEAR_CONDITIONAL_FORMATTING"
  }
}
```

## SET_AUTOFILTER

```json
{
  "stepId": "set-autofilter",
  "target": {
    "type": "BY_RANGE",
    "sheetName": "Sheet1",
    "range": "A1:C20"
  },
  "action": {
    "type": "SET_AUTOFILTER",
    "criteria": [
      {
        "columnId": 2,
        "showButton": true,
        "criterion": {
          "type": "VALUES",
          "values": [
            "Queued",
            "Done"
          ],
          "includeBlank": false
        }
      }
    ],
    "sortState": {
      "range": "A2:C20",
      "caseSensitive": false,
      "columnSort": false,
      "sortMethod": "",
      "conditions": [
        {
          "range": "C2:C20",
          "descending": true,
          "sortBy": ""
        }
      ]
    }
  }
}
```

The range must include a nonblank header row and must not overlap any existing table range.
`criteria` defaults to `[]`; `sortState` is optional. `columnId` is zero-based within the
autofilter range.

## CLEAR_AUTOFILTER

```json
{
  "stepId": "clear-autofilter",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "CLEAR_AUTOFILTER"
  }
}
```

This clears only the sheet-level autofilter. Table-owned autofilters remain attached to their
tables.

## SET_TABLE

```json
{
  "type": "SET_TABLE",
  "table": {
    "name": "DispatchQueue",
    "sheetName": "Sheet1",
    "range": "A1:C20",
    "style": {
      "type": "NONE"
    }
  }
}
{
  "type": "SET_TABLE",
  "table": {
    "name": "DispatchQueue",
    "sheetName": "Sheet1",
    "range": "A1:C20",
    "showTotalsRow": true,
    "hasAutofilter": false,
    "comment": {
      "type": "INLINE",
      "text": "Dispatch tracker"
    },
    "published": true,
    "insertRow": true,
    "insertRowShift": true,
    "headerRowCellStyle": "DispatchHeader",
    "dataCellStyle": "DispatchData",
    "totalsRowCellStyle": "DispatchTotals",
    "style": {
      "type": "NAMED",
      "name": "TableStyleMedium2",
      "showFirstColumn": false,
      "showLastColumn": false,
      "showRowStripes": true,
      "showColumnStripes": false
    },
    "columns": [
      {
        "columnIndex": 0,
        "totalsRowLabel": "Total"
      },
      {
        "columnIndex": 1,
        "totalsRowFunction": "SUM"
      },
      {
        "columnIndex": 2,
        "uniqueName": "status-unique"
      }
    ]
  }
}
```

Table names are workbook-global. Header cells in the first row of the range must be nonblank and
unique case-insensitively. Overlapping different-name tables are rejected. Later value writes and
style patches that touch the table header row keep the table's persisted column metadata aligned
with the visible header cells. `showTotalsRow` is optional and defaults to `false`; when
`totalsRowFunction` is present in a column entry, input is case-insensitive and canonicalized to
Excel's lowercase stored token family. The step target for `SET_TABLE` is
`TableSelector.BY_NAME_ON_SHEET`.

## DELETE_TABLE

```json
{
  "type": "DELETE_TABLE",
  "name": "DispatchQueue",
  "sheetName": "Sheet1"
}
```

`DELETE_TABLE` targets the existing table through `TableSelector.BY_NAME_ON_SHEET`.

## SET_PIVOT_TABLE

```json
{
  "type": "SET_PIVOT_TABLE",
  "pivotTable": {
    "name": "SalesPivot",
    "sheetName": "Report",
    "source": {
      "type": "RANGE",
      "sheetName": "Sheet1",
      "range": "A1:D20"
    },
    "anchor": {
      "topLeftAddress": "C5"
    },
    "rowLabels": [
      "Region"
    ],
    "columnLabels": [
      "Stage"
    ],
    "reportFilters": [],
    "dataFields": [
      {
        "sourceColumnName": "Amount",
        "function": "SUM",
        "displayName": "Total Amount",
        "valueFormat": "#,##0.00"
      }
    ]
  }
}
{
  "type": "SET_PIVOT_TABLE",
  "pivotTable": {
    "name": "NamedPivot",
    "sheetName": "NamedReport",
    "source": {
      "type": "NAMED_RANGE",
      "name": "PivotSource"
    },
    "anchor": {
      "topLeftAddress": "A3"
    },
    "rowLabels": [
      "Region"
    ],
    "columnLabels": [],
    "reportFilters": [
      "Owner"
    ],
    "dataFields": [
      {
        "sourceColumnName": "Amount",
        "function": "SUM"
      }
    ]
  }
}
{
  "type": "SET_PIVOT_TABLE",
  "pivotTable": {
    "name": "TablePivot",
    "sheetName": "TableReport",
    "source": {
      "type": "TABLE",
      "name": "DispatchQueue"
    },
    "anchor": {
      "topLeftAddress": "F4"
    },
    "rowLabels": [
      "Stage"
    ],
    "columnLabels": [],
    "reportFilters": [],
    "dataFields": [
      {
        "sourceColumnName": "Amount",
        "function": "SUM",
        "displayName": "Total Amount"
      }
    ]
  }
}
```

Supported authored pivot sources are `RANGE`, `NAMED_RANGE`, and `TABLE`. `rowLabels`,
`columnLabels`, `reportFilters`, and `dataFields` must use disjoint source columns because POI
persists only one role per pivot field. When `reportFilters` is non-empty,
`anchor.topLeftAddress` must be on Excel row `3` or lower so the page-filter layout has room
above the rendered body. Supported `dataFields[*].function` values are `SUM`, `COUNT`,
`COUNT_NUMS`, `AVERAGE`, `MAX`, `MIN`, `PRODUCT`, `STD_DEV`, `STD_DEVP`, `VAR`, and `VARP`.
The step target for `SET_PIVOT_TABLE` is `PivotTableSelector.BY_NAME_ON_SHEET`.

## DELETE_PIVOT_TABLE

```json
{
  "type": "DELETE_PIVOT_TABLE",
  "name": "SalesPivot",
  "sheetName": "Report"
}
```

`DELETE_PIVOT_TABLE` targets the existing pivot through `PivotTableSelector.BY_NAME_ON_SHEET`.

## APPEND_ROW

```json
{
  "stepId": "append-row",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "APPEND_ROW",
    "values": [
      {
        "type": "TEXT",
        "source": {
          "type": "INLINE",
          "text": "Row label"
        }
      },
      {
        "type": "NUMBER",
        "number": 100
      }
    ]
  }
}
```

Metadata-only blank rows do not affect the append position.
If the append lands on cells that already exist only because of style, hyperlink, or comment
state, that existing state is preserved.

## AUTO_SIZE_COLUMNS

```json
{
  "stepId": "auto-size-columns",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "action": {
    "type": "AUTO_SIZE_COLUMNS"
  }
}
```

Sizing is deterministic and content-based, so headless and Docker runs match local runs.

## SET_NAMED_RANGE

```json
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetTotal",
  "scope": {
    "type": "WORKBOOK"
  },
  "target": {
    "sheetName": "Budget",
    "range": "B4"
  }
}
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetTable",
  "scope": {
    "type": "SHEET",
    "sheetName": "Budget"
  },
  "target": {
    "sheetName": "Budget",
    "range": "A1:B4"
  }
}
{
  "type": "SET_NAMED_RANGE",
  "name": "BudgetRollup",
  "scope": {
    "type": "WORKBOOK"
  },
  "target": {
    "formula": "SUM(Budget!$B$2:$B$5)"
  }
}
```

Named-range targets are normalized to top-left:`bottom-right` order. For example, `"B4:A1"` is
accepted and stored as `"A1:B4"`. Formula-defined targets must set `formula` only.

## DELETE_NAMED_RANGE

```json
{
  "type": "DELETE_NAMED_RANGE",
  "name": "BudgetTotal",
  "scope": {
    "type": "WORKBOOK"
  }
}
```

`SET_COLUMN_WIDTH.widthCharacters` is converted to POI width units with `round(widthCharacters * 256)`. Must be > 0 and ≤ 255.0.
`SET_ROW_HEIGHT.heightPoints` uses Excel point units. Must be > 0 and ≤ 1,638.35 (32,767 twips). `UNMERGE_CELLS` requires an exact merged-range match.
`DELETE_SHEET` returns `INVALID_REQUEST` when the sheet to delete is the only remaining sheet or the last visible sheet in the workbook.

## execution.calculation

Evaluate every formula:

```json
{
  "execution": {
    "calculation": {
      "strategy": {
        "type": "EVALUATE_ALL"
      }
    }
  }
}
```

Evaluate explicit targets only:

```json
{
  "execution": {
    "calculation": {
      "strategy": {
        "type": "EVALUATE_TARGETS",
        "cells": [
          {
            "sheetName": "Budget",
            "address": "D2"
          },
          {
            "sheetName": "Budget",
            "address": "E2"
          }
        ]
      }
    }
  }
}
```

Clear persisted caches:

```json
{
  "execution": {
    "calculation": {
      "strategy": {
        "type": "CLEAR_CACHES_ONLY"
      }
    }
  }
}
```

Mark recalc-on-open without immediate evaluation:

```json
{
  "execution": {
    "calculation": {
      "strategy": {
        "type": "DO_NOT_CALCULATE"
      },
      "markRecalculateOnOpen": true
    }
  }
}
```

---

## Assertion Steps

Assertion steps are ordered alongside mutations and inspections. Passed assertions are echoed in
the success response `assertions[]` array. Failed assertions stop the workflow with
`ASSERTION_FAILED` plus `problem.assertionFailure`. `EXPECT_PRESENT` and `EXPECT_ABSENT` count
matched workbook entities; an exact selector miss is treated as zero observed entities rather than
as a raw `*_NOT_FOUND` read failure.

```json
{
  "stepId": "assert-title",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Budget",
    "address": "A1"
  },
  "assertion": {
    "type": "EXPECT_CELL_VALUE",
    "expectedValue": {
      "type": "TEXT",
      "text": "Quarterly Budget"
    }
  }
}
```

```json
{
  "stepId": "assert-health",
  "target": {
    "type": "BY_NAME",
    "sheetName": "Budget"
  },
  "assertion": {
    "type": "EXPECT_ANALYSIS_MAX_SEVERITY",
    "query": {
      "type": "ANALYZE_FORMULA_HEALTH"
    },
    "maximumSeverity": "INFO"
  }
}
```

```json
{
  "stepId": "assert-no-bad-links",
  "target": {
    "type": "CURRENT"
  },
  "assertion": {
    "type": "EXPECT_ANALYSIS_FINDING_ABSENT",
    "query": {
      "type": "ANALYZE_WORKBOOK_FINDINGS"
    },
    "code": "HYPERLINK_MALFORMED_TARGET"
  }
}
```

```json
{
  "stepId": "assert-budget-state",
  "target": {
    "type": "BY_ADDRESS",
    "sheetName": "Budget",
    "address": "B2"
  },
  "assertion": {
    "type": "ALL_OF",
    "assertions": [
      {
        "type": "EXPECT_CELL_VALUE",
        "expectedValue": {
          "type": "NUMBER",
          "number": 1200
        }
      },
      {
        "type": "NOT",
        "assertion": {
          "type": "EXPECT_DISPLAY_VALUE",
          "displayValue": "0"
        }
      }
    ]
  }
}
```

For a complete mutate-then-verify example, see
[`examples/assertion-request.json`](../examples/assertion-request.json).
For the full assertion-type matrix, selector families, and the less-common fact assertions, see
[OPERATIONS.md](./OPERATIONS.md#assertions).

---

## Inspection Steps

Each inspection is a standalone object for use inside `steps[]`. Every inspection must have a
caller-defined `stepId`.

## GET_WORKBOOK_SUMMARY

```json
{
  "stepId": "workbook",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "GET_WORKBOOK_SUMMARY"
  }
}
```

## GET_PACKAGE_SECURITY

```json
{
  "stepId": "security",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "GET_PACKAGE_SECURITY"
  }
}
```

Returns factual OOXML package-encryption and OOXML package-signature state for the currently open
workbook. This read requires the full-XSSF path; `EVENT_READ` rejects it.

## GET_WORKBOOK_PROTECTION

```json
{
  "stepId": "workbook-protection",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "GET_WORKBOOK_PROTECTION"
  }
}
```

Returns `structureLocked`, `windowsLocked`, `revisionsLocked`, and whether workbook or revisions
password hashes are present.

## GET_CUSTOM_XML_MAPPINGS

```json
{
  "stepId": "custom-xml-mappings",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "GET_CUSTOM_XML_MAPPINGS"
  }
}
```

Returns workbook custom-XML mapping metadata, including identifiers, schema metadata, linked
cells, linked tables, and optional data-binding facts.

## EXPORT_CUSTOM_XML_MAPPING

```json
{
  "stepId": "custom-xml-export",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "EXPORT_CUSTOM_XML_MAPPING",
    "mapping": {
      "mapId": 1,
      "name": "CORSO_mapping"
    },
    "validateSchema": true,
    "encoding": "UTF-8"
  }
}
```

Returns the serialized XML payload for one existing custom-XML mapping together with the factual
mapping metadata used for the export.

## GET_NAMED_RANGES

```json
{
  "stepId": "named-ranges",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "GET_NAMED_RANGES"
  }
}
{
  "stepId": "selected-named-ranges",
  "target": {
    "type": "ANY_OF",
    "selectors": [
      {
        "type": "WORKBOOK_SCOPE",
        "name": "BudgetTotal"
      },
      {
        "type": "SHEET_SCOPE",
        "sheetName": "Budget",
        "name": "BudgetTable"
      }
    ]
  },
  "query": {
    "type": "GET_NAMED_RANGES"
  }
}
```

## GET_SHEET_SUMMARY

```json
{
  "stepId": "sheet-summary",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "query": {
    "type": "GET_SHEET_SUMMARY"
  }
}
```

## GET_CELLS

```json
{
  "stepId": "cells",
  "target": {
    "type": "BY_ADDRESSES",
    "sheetName": "Sheet1",
    "addresses": [
      "A1",
      "B4",
      "C10"
    ]
  },
  "query": {
    "type": "GET_CELLS"
  }
}
```

String cells return `stringValue` and, when the stored cell contains authored rich text, an
optional ordered `richText` run list with effective per-run font facts.
Read-side style colors are structured objects with `rgb` plus optional `theme`, `indexed`, and
`tint`; gradient fills appear under `style.fill.gradient`.

## GET_ARRAY_FORMULAS

```json
{
  "stepId": "array-formulas",
  "target": {
    "type": "BY_NAME",
    "name": "Calc"
  },
  "query": {
    "type": "GET_ARRAY_FORMULAS"
  }
}
```

Returns factual array-formula group metadata for the selected sheets, including the stored range,
top-left address, normalized formula text, and whether the group is single-cell.

## GET_WINDOW

```json
{
  "stepId": "window",
  "target": {
    "type": "RECTANGULAR_WINDOW",
    "sheetName": "Sheet1",
    "topLeftAddress": "A1",
    "rowCount": 5,
    "columnCount": 4
  },
  "query": {
    "type": "GET_WINDOW"
  }
}
```

`rowCount * columnCount` must not exceed 250,000.

## GET_MERGED_REGIONS

```json
{
  "stepId": "merged",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "query": {
    "type": "GET_MERGED_REGIONS"
  }
}
```

## GET_HYPERLINKS

Response hyperlinks reuse the same discriminated shape as `SET_HYPERLINK` targets. `FILE`
responses use the `path` field with a normalized plain path string.

```json
{
  "stepId": "hyperlinks",
  "target": {
    "type": "ALL_USED_IN_SHEET",
    "sheetName": "Sheet1"
  },
  "query": {
    "type": "GET_HYPERLINKS"
  }
}
{
  "stepId": "selected-hyperlinks",
  "target": {
    "type": "BY_ADDRESSES",
    "sheetName": "Sheet1",
    "addresses": [
      "A1",
      "B4"
    ]
  },
  "query": {
    "type": "GET_HYPERLINKS"
  }
}
```

## GET_COMMENTS

```json
{
  "stepId": "comments",
  "target": {
    "type": "ALL_USED_IN_SHEET",
    "sheetName": "Sheet1"
  },
  "query": {
    "type": "GET_COMMENTS"
  }
}
{
  "stepId": "selected-comments",
  "target": {
    "type": "BY_ADDRESSES",
    "sheetName": "Sheet1",
    "addresses": [
      "A1",
      "B4"
    ]
  },
  "query": {
    "type": "GET_COMMENTS"
  }
}
```

Returned comments can include ordered rich-text `runs` plus an `anchor` with zero-based
`firstColumn`, `firstRow`, `lastColumn`, and `lastRow` bounds.

## GET_DRAWING_OBJECTS

```json
{
  "stepId": "drawing-objects",
  "target": {
    "type": "ALL_ON_SHEET",
    "sheetName": "Ops"
  },
  "query": {
    "type": "GET_DRAWING_OBJECTS"
  }
}
```

Returned entries are `PICTURE`, `CHART`, `SHAPE`, `EMBEDDED_OBJECT`, or `SIGNATURE_LINE`.
`CHART` entries expose `supported`, ordered `plotTypeTokens`, and title text. `SIGNATURE_LINE`
entries expose signer metadata, comment-permission state, optional preview-image facts, and the
stored anchor. Read anchors can be `TWO_CELL`,
`ONE_CELL`, or `ABSOLUTE`.

## GET_CHARTS

```json
{
  "stepId": "charts",
  "target": {
    "type": "ALL_ON_SHEET",
    "sheetName": "Ops"
  },
  "query": {
    "type": "GET_CHARTS"
  }
}
```

Returned chart entries expose chart-level fields plus ordered `plots`. Plot entries are
`AREA`, `AREA_3D`, `BAR`, `BAR_3D`, `DOUGHNUT`, `LINE`, `LINE_3D`, `PIE`, `PIE_3D`, `RADAR`,
`SCATTER`, `SURFACE`, `SURFACE_3D`, or `UNSUPPORTED`. Unsupported loaded plots preserve
`plotTypeTokens` and `detail`.
Blank loaded chart titles normalize to `NONE`. Sparse literal caches surface missing positions as
empty strings. If a chart relation is gone but the graphic frame still exists, `GET_CHARTS` skips
the broken chart and `GET_DRAWING_OBJECTS` reports the surviving frame as read-only
`GRAPHIC_FRAME`.

## GET_DRAWING_OBJECT_PAYLOAD

```json
{
  "stepId": "picture-payload",
  "target": {
    "type": "BY_NAME",
    "sheetName": "Ops",
    "objectName": "OpsPicture"
  },
  "query": {
    "type": "GET_DRAWING_OBJECT_PAYLOAD"
  }
}
{
  "stepId": "embedded-payload",
  "target": {
    "type": "BY_NAME",
    "sheetName": "Ops",
    "objectName": "OpsEmbed"
  },
  "query": {
    "type": "GET_DRAWING_OBJECT_PAYLOAD"
  }
}
```

Payload extraction is only for named pictures and embedded objects. Non-binary drawing objects
such as signature lines, connectors, and simple shapes are rejected.

## GET_SHEET_LAYOUT

```json
{
  "stepId": "layout",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "query": {
    "type": "GET_SHEET_LAYOUT"
  }
}
```

Row and column entries report explicit size plus `hidden`, `outlineLevel`, and `collapsed`
where Excel exposes that state.

`layout.presentation` reports screen display flags, right-to-left layout, tab color,
outline-summary placement, sheet-default sizing, and ignored-error suppression.

## GET_PRINT_LAYOUT

```json
{
  "stepId": "print-layout",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "query": {
    "type": "GET_PRINT_LAYOUT"
  }
}
```

`printLayout.setup` carries advanced page-setup facts such as margins, print-gridline output,
copies, first-page numbering, and explicit row or column breaks.

## GET_DATA_VALIDATIONS

```json
{
  "stepId": "data-validations",
  "target": {
    "type": "ALL_ON_SHEET",
    "sheetName": "Sheet1"
  },
  "query": {
    "type": "GET_DATA_VALIDATIONS"
  }
}
{
  "stepId": "selected-data-validations",
  "target": {
    "type": "BY_RANGES",
    "sheetName": "Sheet1",
    "ranges": [
      "B2:B20",
      "C2:C20"
    ]
  },
  "query": {
    "type": "GET_DATA_VALIDATIONS"
  }
}
```

## GET_AUTOFILTERS

```json
{
  "stepId": "autofilters",
  "target": {
    "type": "BY_NAME",
    "name": "Sheet1"
  },
  "query": {
    "type": "GET_AUTOFILTERS"
  }
}
```

Entries are `SHEET` or `TABLE` and can include persisted `filterColumns` plus `sortState`.
Persisted sort-state metadata is reported exactly as stored, including blank raw ranges.

## GET_CONDITIONAL_FORMATTING

```json
{
  "stepId": "conditional-formatting",
  "target": {
    "type": "ALL_ON_SHEET",
    "sheetName": "Sheet1"
  },
  "query": {
    "type": "GET_CONDITIONAL_FORMATTING"
  }
}
{
  "stepId": "selected-conditional-formatting",
  "target": {
    "type": "BY_RANGES",
    "sheetName": "Sheet1",
    "ranges": [
      "A2:D20",
      "F2:F20"
    ]
  },
  "query": {
    "type": "GET_CONDITIONAL_FORMATTING"
  }
}
```

Read entries may report `FORMULA_RULE`, `CELL_VALUE_RULE`, `COLOR_SCALE_RULE`, `DATA_BAR_RULE`,
`ICON_SET_RULE`, or `UNSUPPORTED_RULE`.
Malformed loaded rule metadata degrades to `UNSUPPORTED_RULE` instead of aborting the read.

## GET_TABLES

```json
{
  "stepId": "tables",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "GET_TABLES"
  }
}
{
  "stepId": "selected-tables",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "DispatchQueue",
      "Trips"
    ]
  },
  "query": {
    "type": "GET_TABLES"
  }
}
```

Returned tables include persisted `columns` metadata plus advanced flags and style names such as
`comment`, `published`, `insertRow`, `insertRowShift`, `headerRowCellStyle`, `dataCellStyle`, and
`totalsRowCellStyle`.

## GET_PIVOT_TABLES

```json
{
  "stepId": "pivots",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "GET_PIVOT_TABLES"
  }
}
{
  "stepId": "selected-pivots",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "SalesPivot",
      "NamedPivot"
    ]
  },
  "query": {
    "type": "GET_PIVOT_TABLES"
  }
}
```

Returned entries are `SUPPORTED` or `UNSUPPORTED`. Supported pivot reports include factual `source`,
stored `anchor`, `rowLabels`, `columnLabels`, `reportFilters`, `dataFields`, and
`valuesAxisOnColumns`. Unsupported or malformed loaded pivots are preserved as explicit
`UNSUPPORTED` entries with a human-readable `detail` instead of aborting the read.

## GET_FORMULA_SURFACE

```json
{
  "stepId": "formula-surface",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "GET_FORMULA_SURFACE"
  }
}
{
  "stepId": "selected-formula-surface",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "Sheet1",
      "Sheet2"
    ]
  },
  "query": {
    "type": "GET_FORMULA_SURFACE"
  }
}
```

Response shape: `analysis.totalFormulaCellCount` plus `analysis.sheets[*]` entries with
`sheetName`, `formulaCellCount`, `distinctFormulaCount`, and grouped `formulas[*]`
(`formula`, `occurrenceCount`, `addresses`).

## GET_SHEET_SCHEMA

```json
{
  "stepId": "schema",
  "target": {
    "type": "RECTANGULAR_WINDOW",
    "sheetName": "Sheet1",
    "topLeftAddress": "A1",
    "rowCount": 5,
    "columnCount": 4
  },
  "query": {
    "type": "GET_SHEET_SCHEMA"
  }
}
```

`rowCount * columnCount` must not exceed 250,000. When the first row is entirely blank,
`dataRowCount` in the response is `0`.

## GET_NAMED_RANGE_SURFACE

```json
{
  "stepId": "named-range-surface",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "GET_NAMED_RANGE_SURFACE"
  }
}
{
  "stepId": "selected-named-range-surface",
  "target": {
    "type": "ANY_OF",
    "selectors": [
      {
        "type": "WORKBOOK_SCOPE",
        "name": "BudgetTotal"
      },
      {
        "type": "SHEET_SCOPE",
        "sheetName": "Budget",
        "name": "BudgetTable"
      }
    ]
  },
  "query": {
    "type": "GET_NAMED_RANGE_SURFACE"
  }
}
```

Response shape: `analysis.workbookScopedCount`, `sheetScopedCount`, `rangeBackedCount`,
`formulaBackedCount`, and `namedRanges[*]` entries with `name`, `scope`, `refersToFormula`,
and `kind`.

## ANALYZE_FORMULA_HEALTH

```json
{
  "stepId": "formula-health",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "ANALYZE_FORMULA_HEALTH"
  }
}
```

Response shape: `analysis.checkedFormulaCellCount`, severity `summary`, and `findings`.

## ANALYZE_DATA_VALIDATION_HEALTH

```json
{
  "stepId": "data-validation-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "Sheet1",
      "Sheet2"
    ]
  },
  "query": {
    "type": "ANALYZE_DATA_VALIDATION_HEALTH"
  }
}
```

## ANALYZE_CONDITIONAL_FORMATTING_HEALTH

```json
{
  "stepId": "conditional-formatting-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "Sheet1",
      "Sheet2"
    ]
  },
  "query": {
    "type": "ANALYZE_CONDITIONAL_FORMATTING_HEALTH"
  }
}
```

## ANALYZE_AUTOFILTER_HEALTH

```json
{
  "stepId": "autofilter-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "Sheet1",
      "Sheet2"
    ]
  },
  "query": {
    "type": "ANALYZE_AUTOFILTER_HEALTH"
  }
}
```

## ANALYZE_TABLE_HEALTH

```json
{
  "stepId": "table-health",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "ANALYZE_TABLE_HEALTH"
  }
}
{
  "stepId": "selected-table-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "DispatchQueue",
      "Trips"
    ]
  },
  "query": {
    "type": "ANALYZE_TABLE_HEALTH"
  }
}
```

## ANALYZE_PIVOT_TABLE_HEALTH

```json
{
  "stepId": "pivot-health",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "ANALYZE_PIVOT_TABLE_HEALTH"
  }
}
{
  "stepId": "selected-pivot-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "SalesPivot",
      "NamedPivot"
    ]
  },
  "query": {
    "type": "ANALYZE_PIVOT_TABLE_HEALTH"
  }
}
```

Pivot health reports findings such as missing cache parts, missing workbook cache definitions,
broken sources, duplicate names, synthetic fallback names, or unsupported persisted detail.

## ANALYZE_HYPERLINK_HEALTH

```json
{
  "stepId": "hyperlink-health",
  "target": {
    "type": "BY_NAMES",
    "names": [
      "Sheet1",
      "Sheet2"
    ]
  },
  "query": {
    "type": "ANALYZE_HYPERLINK_HEALTH"
  }
}
```

Relative `FILE` targets are resolved against the workbook's persisted path when one exists.
Unsaved workbooks report unresolved relative file targets instead of treating them as healthy.

## ANALYZE_NAMED_RANGE_HEALTH

```json
{
  "stepId": "named-range-health",
  "target": {
    "type": "ALL"
  },
  "query": {
    "type": "ANALYZE_NAMED_RANGE_HEALTH"
  }
}
```

Response shape: `analysis.checkedNamedRangeCount`, severity `summary`, and `findings`.

## ANALYZE_WORKBOOK_FINDINGS

```json
{
  "stepId": "workbook-findings",
  "target": {
    "type": "CURRENT"
  },
  "query": {
    "type": "ANALYZE_WORKBOOK_FINDINGS"
  }
}
```

Response shape: `analysis.summary` plus one flat `analysis.findings` list.

`ANALYZE_WORKBOOK_FINDINGS` aggregates every shipped health family: formula, data validation,
conditional formatting, autofilter, table, pivot table, hyperlink, and named range. It is the
primary workbook-health check and pairs naturally with `persistence.type=NONE`.

Selector snippets:

```json
{
  "type": "ALL_USED_IN_SHEET",
  "sheetName": "Sheet1"
}
{
  "type": "BY_ADDRESSES",
  "sheetName": "Sheet1",
  "addresses": [
    "A1",
    "B4",
    "C10"
  ]
}
```

```json
{
  "type": "ALL_ON_SHEET",
  "sheetName": "Sheet1"
}
{
  "type": "BY_NAMES",
  "names": [
    "Sheet1",
    "Sheet2"
  ]
}
```

```json
{
  "type": "ALL"
}
{
  "type": "BY_NAMES",
  "names": [
    "SalesPivot",
    "NamedPivot"
  ]
}
```

```json
{
  "type": "ALL"
}
{
  "type": "ANY_OF",
  "selectors": [
    {
      "type": "WORKBOOK_SCOPE",
      "name": "BudgetTotal"
    },
    {
      "type": "SHEET_SCOPE",
      "sheetName": "Budget",
      "name": "BudgetTable"
    }
  ]
}
```

```json
{
  "type": "ALL"
}
{
  "type": "BY_NAMES",
  "names": [
    "DispatchQueue",
    "Trips"
  ]
}
```
```
