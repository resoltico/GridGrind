<!-- Generated from core metadata; do not edit. -->
# mutationActionTypes

## ENSURE_SHEET

Create the sheet if it does not already exist.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "ENSURE_SHEET",
  "summary" : "Create the sheet if it does not already exist.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-ensure-sheet",
      "action" : {
        "type" : "ENSURE_SHEET"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## RENAME_SHEET

Rename an existing sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| newSheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "RENAME_SHEET",
  "summary" : "Rename an existing sheet.",
  "fields" : [ {
    "name" : "newSheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-rename-sheet",
      "action" : {
        "type" : "RENAME_SHEET",
        "newSheetName" : "sample-newSheetName"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## DELETE_SHEET

Delete an existing sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "DELETE_SHEET",
  "summary" : "Delete an existing sheet.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-delete-sheet",
      "action" : {
        "type" : "DELETE_SHEET"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## MOVE_SHEET

Move a sheet to a zero-based workbook position.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| targetIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "MOVE_SHEET",
  "summary" : "Move a sheet to a zero-based workbook position.",
  "fields" : [ {
    "name" : "targetIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-move-sheet",
      "action" : {
        "type" : "MOVE_SHEET",
        "targetIndex" : 1
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## COPY_SHEET

Copy one sheet into a new visible, unselected sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| newSheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| position | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"sheetCopyPositionTypes"} | {type=APPEND_AT_END} |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "COPY_SHEET",
  "summary" : "Copy one sheet into a new visible, unselected sheet.",
  "fields" : [ {
    "name" : "newSheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "position",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "sheetCopyPositionTypes"
    },
    "enumValues" : [ ],
    "defaultValue" : {
      "type" : "APPEND_AT_END"
    }
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-copy-sheet",
      "action" : {
        "type" : "COPY_SHEET",
        "newSheetName" : "sample-newSheetName"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_ACTIVE_SHEET

Set the active sheet and ensure it is selected.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_ACTIVE_SHEET",
  "summary" : "Set the active sheet and ensure it is selected.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-active-sheet",
      "action" : {
        "type" : "SET_ACTIVE_SHEET"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_SELECTED_SHEETS

Set the selected visible sheet set.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAMES]]].

```json
{
  "id" : "SET_SELECTED_SHEETS",
  "summary" : "Set the selected visible sheet set.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAMES" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-selected-sheets",
      "action" : {
        "type" : "SET_SELECTED_SHEETS"
      },
      "target" : {
        "type" : "SHEET_BY_NAMES",
        "names" : [ "sample-names" ]
      }
    },
    "notes" : [ ]
  }
}
```

## SET_SHEET_VISIBILITY

Set one sheet visibility state.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| visibility | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_SHEET_VISIBILITY",
  "summary" : "Set one sheet visibility state.",
  "fields" : [ {
    "name" : "visibility",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "VISIBLE", "HIDDEN", "VERY_HIDDEN" ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-sheet-visibility",
      "action" : {
        "type" : "SET_SHEET_VISIBILITY",
        "visibility" : "VISIBLE"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_SHEET_PROTECTION

Enable sheet protection with the exact supported lock flags.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| protection | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"sheetProtectionSettingsType"} |  |
| passwordRef | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"secretReferenceType"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_SHEET_PROTECTION",
  "summary" : "Enable sheet protection with the exact supported lock flags.",
  "fields" : [ {
    "name" : "protection",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "sheetProtectionSettingsType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "passwordRef",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "secretReferenceType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-sheet-protection",
      "action" : {
        "type" : "SET_SHEET_PROTECTION",
        "protection" : {
          "autoFilterLocked" : false,
          "deleteColumnsLocked" : false,
          "deleteRowsLocked" : false,
          "formatCellsLocked" : false,
          "formatColumnsLocked" : false,
          "formatRowsLocked" : false,
          "insertColumnsLocked" : false,
          "insertHyperlinksLocked" : false,
          "insertRowsLocked" : false,
          "objectsLocked" : false,
          "pivotTablesLocked" : false,
          "scenariosLocked" : false,
          "selectLockedCellsLocked" : false,
          "selectUnlockedCellsLocked" : false,
          "sortLocked" : false
        }
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## CLEAR_SHEET_PROTECTION

Disable sheet protection entirely.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "CLEAR_SHEET_PROTECTION",
  "summary" : "Disable sheet protection entirely.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-clear-sheet-protection",
      "action" : {
        "type" : "CLEAR_SHEET_PROTECTION"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_WORKBOOK_PROTECTION

Enable workbook-level protection and optional password hashes.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| protection | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"workbookProtectionSpecType"} |  |

Targets: [TargetSelectorEntry[family=WorkbookSelector, typeIds=[WORKBOOK_CURRENT]]].

```json
{
  "id" : "SET_WORKBOOK_PROTECTION",
  "summary" : "Enable workbook-level protection and optional password hashes.",
  "fields" : [ {
    "name" : "protection",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "workbookProtectionSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "WorkbookSelector",
    "typeIds" : [ "WORKBOOK_CURRENT" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-workbook-protection",
      "action" : {
        "type" : "SET_WORKBOOK_PROTECTION",
        "protection" : { }
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ ]
  }
}
```

## CLEAR_WORKBOOK_PROTECTION

Clear workbook-level protection and stored workbook passwords.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=WorkbookSelector, typeIds=[WORKBOOK_CURRENT]]].

```json
{
  "id" : "CLEAR_WORKBOOK_PROTECTION",
  "summary" : "Clear workbook-level protection and stored workbook passwords.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "WorkbookSelector",
    "typeIds" : [ "WORKBOOK_CURRENT" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-clear-workbook-protection",
      "action" : {
        "type" : "CLEAR_WORKBOOK_PROTECTION"
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ ]
  }
}
```

## MERGE_CELLS

Merge a rectangular A1-style range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_BY_RANGE]]].

```json
{
  "id" : "MERGE_CELLS",
  "summary" : "Merge a rectangular A1-style range.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_BY_RANGE" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-merge-cells",
      "action" : {
        "type" : "MERGE_CELLS"
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ ]
  }
}
```

## UNMERGE_CELLS

Remove one merged region by exact range match.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_BY_RANGE]]].

```json
{
  "id" : "UNMERGE_CELLS",
  "summary" : "Remove one merged region by exact range match.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_BY_RANGE" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-unmerge-cells",
      "action" : {
        "type" : "UNMERGE_CELLS"
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_COLUMN_WIDTH

Set one or more column widths in Excel character units.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| widthCharacters | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

Targets: [TargetSelectorEntry[family=ColumnBandSelector, typeIds=[COLUMN_BAND_SPAN]]].

```json
{
  "id" : "SET_COLUMN_WIDTH",
  "summary" : "Set one or more column widths in Excel character units.",
  "fields" : [ {
    "name" : "widthCharacters",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "ColumnBandSelector",
    "typeIds" : [ "COLUMN_BAND_SPAN" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-column-width",
      "action" : {
        "type" : "SET_COLUMN_WIDTH",
        "widthCharacters" : 1
      },
      "target" : {
        "type" : "COLUMN_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstColumnIndex" : 0,
        "lastColumnIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## SET_ROW_HEIGHT

Set one or more row heights in Excel point units.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| heightPoints | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

Targets: [TargetSelectorEntry[family=RowBandSelector, typeIds=[ROW_BAND_SPAN]]].

```json
{
  "id" : "SET_ROW_HEIGHT",
  "summary" : "Set one or more row heights in Excel point units.",
  "fields" : [ {
    "name" : "heightPoints",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "RowBandSelector",
    "typeIds" : [ "ROW_BAND_SPAN" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-row-height",
      "action" : {
        "type" : "SET_ROW_HEIGHT",
        "heightPoints" : 1
      },
      "target" : {
        "type" : "ROW_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstRowIndex" : 0,
        "lastRowIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## INSERT_ROWS

Insert one or more blank rows before rowIndex.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=RowBandSelector, typeIds=[ROW_BAND_INSERTION]]].

```json
{
  "id" : "INSERT_ROWS",
  "summary" : "Insert one or more blank rows before rowIndex.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "RowBandSelector",
    "typeIds" : [ "ROW_BAND_INSERTION" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-insert-rows",
      "action" : {
        "type" : "INSERT_ROWS"
      },
      "target" : {
        "type" : "ROW_BAND_INSERTION",
        "sheetName" : "Sheet1",
        "beforeRowIndex" : 1,
        "rowCount" : 1
      }
    },
    "notes" : [ ]
  }
}
```

## DELETE_ROWS

Delete one inclusive zero-based row band.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=RowBandSelector, typeIds=[ROW_BAND_SPAN]]].

```json
{
  "id" : "DELETE_ROWS",
  "summary" : "Delete one inclusive zero-based row band.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "RowBandSelector",
    "typeIds" : [ "ROW_BAND_SPAN" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-delete-rows",
      "action" : {
        "type" : "DELETE_ROWS"
      },
      "target" : {
        "type" : "ROW_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstRowIndex" : 0,
        "lastRowIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## SHIFT_ROWS

Move one inclusive zero-based row band by delta rows.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| delta | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

Targets: [TargetSelectorEntry[family=RowBandSelector, typeIds=[ROW_BAND_SPAN]]].

```json
{
  "id" : "SHIFT_ROWS",
  "summary" : "Move one inclusive zero-based row band by delta rows.",
  "fields" : [ {
    "name" : "delta",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "RowBandSelector",
    "typeIds" : [ "ROW_BAND_SPAN" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-shift-rows",
      "action" : {
        "type" : "SHIFT_ROWS",
        "delta" : 1
      },
      "target" : {
        "type" : "ROW_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstRowIndex" : 0,
        "lastRowIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## INSERT_COLUMNS

Insert one or more blank columns before columnIndex.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=ColumnBandSelector, typeIds=[COLUMN_BAND_INSERTION]]].

Preconditions: [ColumnEditsBeforeFormulaAuthoring[]].

```json
{
  "id" : "INSERT_COLUMNS",
  "summary" : "Insert one or more blank columns before columnIndex.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "ColumnBandSelector",
    "typeIds" : [ "COLUMN_BAND_INSERTION" ]
  } ],
  "preconditions" : [ {
    "type" : "NO_COLUMN_EDITS_AFTER_FORMULA_AUTHORING"
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-insert-columns",
      "action" : {
        "type" : "INSERT_COLUMNS"
      },
      "target" : {
        "type" : "COLUMN_BAND_INSERTION",
        "sheetName" : "Sheet1",
        "beforeColumnIndex" : 1,
        "columnCount" : 1
      }
    },
    "notes" : [ ]
  }
}
```

## DELETE_COLUMNS

Delete one inclusive zero-based column band.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=ColumnBandSelector, typeIds=[COLUMN_BAND_SPAN]]].

Preconditions: [ColumnEditsBeforeFormulaAuthoring[]].

```json
{
  "id" : "DELETE_COLUMNS",
  "summary" : "Delete one inclusive zero-based column band.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "ColumnBandSelector",
    "typeIds" : [ "COLUMN_BAND_SPAN" ]
  } ],
  "preconditions" : [ {
    "type" : "NO_COLUMN_EDITS_AFTER_FORMULA_AUTHORING"
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-delete-columns",
      "action" : {
        "type" : "DELETE_COLUMNS"
      },
      "target" : {
        "type" : "COLUMN_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstColumnIndex" : 0,
        "lastColumnIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## SHIFT_COLUMNS

Move one inclusive zero-based column band by delta columns.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| delta | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

Targets: [TargetSelectorEntry[family=ColumnBandSelector, typeIds=[COLUMN_BAND_SPAN]]].

Preconditions: [ColumnEditsBeforeFormulaAuthoring[]].

```json
{
  "id" : "SHIFT_COLUMNS",
  "summary" : "Move one inclusive zero-based column band by delta columns.",
  "fields" : [ {
    "name" : "delta",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "ColumnBandSelector",
    "typeIds" : [ "COLUMN_BAND_SPAN" ]
  } ],
  "preconditions" : [ {
    "type" : "NO_COLUMN_EDITS_AFTER_FORMULA_AUTHORING"
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-shift-columns",
      "action" : {
        "type" : "SHIFT_COLUMNS",
        "delta" : 1
      },
      "target" : {
        "type" : "COLUMN_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstColumnIndex" : 0,
        "lastColumnIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## SET_ROW_VISIBILITY

Set the hidden state for one inclusive zero-based row band.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| hidden | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

Targets: [TargetSelectorEntry[family=RowBandSelector, typeIds=[ROW_BAND_SPAN]]].

```json
{
  "id" : "SET_ROW_VISIBILITY",
  "summary" : "Set the hidden state for one inclusive zero-based row band.",
  "fields" : [ {
    "name" : "hidden",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "RowBandSelector",
    "typeIds" : [ "ROW_BAND_SPAN" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-row-visibility",
      "action" : {
        "type" : "SET_ROW_VISIBILITY",
        "hidden" : true
      },
      "target" : {
        "type" : "ROW_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstRowIndex" : 0,
        "lastRowIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## SET_COLUMN_VISIBILITY

Set the hidden state for one inclusive zero-based column band.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| hidden | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

Targets: [TargetSelectorEntry[family=ColumnBandSelector, typeIds=[COLUMN_BAND_SPAN]]].

```json
{
  "id" : "SET_COLUMN_VISIBILITY",
  "summary" : "Set the hidden state for one inclusive zero-based column band.",
  "fields" : [ {
    "name" : "hidden",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "ColumnBandSelector",
    "typeIds" : [ "COLUMN_BAND_SPAN" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-column-visibility",
      "action" : {
        "type" : "SET_COLUMN_VISIBILITY",
        "hidden" : true
      },
      "target" : {
        "type" : "COLUMN_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstColumnIndex" : 0,
        "lastColumnIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## GROUP_ROWS

Apply one outline group to an inclusive zero-based row band.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| collapsed | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

Targets: [TargetSelectorEntry[family=RowBandSelector, typeIds=[ROW_BAND_SPAN]]].

```json
{
  "id" : "GROUP_ROWS",
  "summary" : "Apply one outline group to an inclusive zero-based row band.",
  "fields" : [ {
    "name" : "collapsed",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "RowBandSelector",
    "typeIds" : [ "ROW_BAND_SPAN" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-group-rows",
      "action" : {
        "type" : "GROUP_ROWS",
        "collapsed" : false
      },
      "target" : {
        "type" : "ROW_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstRowIndex" : 0,
        "lastRowIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## UNGROUP_ROWS

Remove outline grouping from one inclusive zero-based row band.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=RowBandSelector, typeIds=[ROW_BAND_SPAN]]].

```json
{
  "id" : "UNGROUP_ROWS",
  "summary" : "Remove outline grouping from one inclusive zero-based row band.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "RowBandSelector",
    "typeIds" : [ "ROW_BAND_SPAN" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-ungroup-rows",
      "action" : {
        "type" : "UNGROUP_ROWS"
      },
      "target" : {
        "type" : "ROW_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstRowIndex" : 0,
        "lastRowIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## GROUP_COLUMNS

Apply one outline group to an inclusive zero-based column band.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| collapsed | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

Targets: [TargetSelectorEntry[family=ColumnBandSelector, typeIds=[COLUMN_BAND_SPAN]]].

```json
{
  "id" : "GROUP_COLUMNS",
  "summary" : "Apply one outline group to an inclusive zero-based column band.",
  "fields" : [ {
    "name" : "collapsed",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "ColumnBandSelector",
    "typeIds" : [ "COLUMN_BAND_SPAN" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-group-columns",
      "action" : {
        "type" : "GROUP_COLUMNS",
        "collapsed" : false
      },
      "target" : {
        "type" : "COLUMN_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstColumnIndex" : 0,
        "lastColumnIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## UNGROUP_COLUMNS

Remove outline grouping from one inclusive zero-based column band.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=ColumnBandSelector, typeIds=[COLUMN_BAND_SPAN]]].

```json
{
  "id" : "UNGROUP_COLUMNS",
  "summary" : "Remove outline grouping from one inclusive zero-based column band.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "ColumnBandSelector",
    "typeIds" : [ "COLUMN_BAND_SPAN" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-ungroup-columns",
      "action" : {
        "type" : "UNGROUP_COLUMNS"
      },
      "target" : {
        "type" : "COLUMN_BAND_SPAN",
        "sheetName" : "Sheet1",
        "firstColumnIndex" : 0,
        "lastColumnIndex" : 0
      }
    },
    "notes" : [ ]
  }
}
```

## SET_SHEET_PANE

Apply one explicit pane state to a sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pane | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"paneSpecTypes"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_SHEET_PANE",
  "summary" : "Apply one explicit pane state to a sheet.",
  "fields" : [ {
    "name" : "pane",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "paneSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-sheet-pane",
      "action" : {
        "type" : "SET_SHEET_PANE",
        "pane" : {
          "type" : "NONE"
        }
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_SHEET_ZOOM

Set the sheet zoom percentage.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| zoomPercent | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_SHEET_ZOOM",
  "summary" : "Set the sheet zoom percentage.",
  "fields" : [ {
    "name" : "zoomPercent",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-sheet-zoom",
      "action" : {
        "type" : "SET_SHEET_ZOOM",
        "zoomPercent" : 100
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_SHEET_PRESENTATION

Apply supported sheet-presentation state such as display flags and defaults.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| presentation | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"sheetPresentationSpecType"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_SHEET_PRESENTATION",
  "summary" : "Apply supported sheet-presentation state such as display flags and defaults.",
  "fields" : [ {
    "name" : "presentation",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "sheetPresentationSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-sheet-presentation",
      "action" : {
        "type" : "SET_SHEET_PRESENTATION",
        "presentation" : {
          "display" : {
            "displayGridlines" : false,
            "displayZeros" : false,
            "displayRowColHeadings" : false,
            "displayFormulas" : false,
            "rightToLeft" : false
          },
          "outlineSummary" : {
            "rowSumsBelow" : false,
            "rowSumsRight" : false
          },
          "sheetDefaults" : {
            "defaultColumnWidth" : 1,
            "defaultRowHeightPoints" : 1
          },
          "ignoredErrors" : [ {
            "range" : "A1:B2",
            "errorTypes" : [ "CALCULATED_COLUMN" ]
          } ]
        }
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_PRINT_LAYOUT

Apply supported print-layout state to a sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| printLayout | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"printLayoutSpecType"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_PRINT_LAYOUT",
  "summary" : "Apply supported print-layout state to a sheet.",
  "fields" : [ {
    "name" : "printLayout",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "printLayoutSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-print-layout",
      "action" : {
        "type" : "SET_PRINT_LAYOUT",
        "printLayout" : {
          "printArea" : {
            "type" : "NONE"
          },
          "orientation" : "PORTRAIT",
          "scaling" : {
            "type" : "AUTOMATIC"
          },
          "repeatingRows" : {
            "type" : "NONE"
          },
          "repeatingColumns" : {
            "type" : "NONE"
          },
          "header" : {
            "left" : {
              "type" : "INLINE",
              "text" : "Sample text"
            },
            "center" : {
              "type" : "INLINE",
              "text" : "Sample text"
            },
            "right" : {
              "type" : "INLINE",
              "text" : "Sample text"
            }
          },
          "footer" : {
            "left" : {
              "type" : "INLINE",
              "text" : "Sample text"
            },
            "center" : {
              "type" : "INLINE",
              "text" : "Sample text"
            },
            "right" : {
              "type" : "INLINE",
              "text" : "Sample text"
            }
          },
          "setup" : {
            "margins" : {
              "left" : 1,
              "right" : 1,
              "top" : 1,
              "bottom" : 1,
              "header" : 1,
              "footer" : 1
            },
            "printGridlines" : false,
            "horizontallyCentered" : false,
            "verticallyCentered" : false,
            "paperSize" : 1,
            "draft" : false,
            "blackAndWhite" : false,
            "copies" : 1,
            "useFirstPageNumber" : false,
            "firstPageNumber" : 1,
            "rowBreaks" : [ 1 ],
            "columnBreaks" : [ 1 ]
          }
        }
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## CLEAR_PRINT_LAYOUT

Clear the supported print-layout state from a sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "CLEAR_PRINT_LAYOUT",
  "summary" : "Clear the supported print-layout state from a sheet.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-clear-print-layout",
      "action" : {
        "type" : "CLEAR_PRINT_LAYOUT"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## AUTO_SIZE_COLUMNS

Auto-size all populated columns on the sheet to fit their content.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "AUTO_SIZE_COLUMNS",
  "summary" : "Auto-size all populated columns on the sheet to fit their content.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS", "WORKBOOK_STRUCTURE", "SHEET_STRUCTURE", "ROWS", "COLUMNS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "WORKBOOK_METADATA", "WORKBOOK_SECURITY" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-auto-size-columns",
      "action" : {
        "type" : "AUTO_SIZE_COLUMNS"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_CELL

Write one typed value to a single cell.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellValueSpecTypes"} |  |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_BY_ADDRESS]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "SET_CELL",
  "summary" : "Write one typed value to a single cell.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellValueSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_BY_ADDRESS" ]
  }, {
    "family" : "TableCellSelector",
    "typeIds" : [ "TABLE_CELL_BY_COLUMN_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-cell",
      "action" : {
        "type" : "SET_CELL",
        "value" : {
          "type" : "TEXT",
          "value" : {
            "type" : "INLINE",
            "text" : "Sample text"
          }
        }
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; SET_CELL accepts CELL_BY_ADDRESS, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## SET_RANGE

Write a rectangular grid of typed or compact homogeneous values.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| rows | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellGridSpecTypes"} |  |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_BY_RANGE]]].

```json
{
  "id" : "SET_RANGE",
  "summary" : "Write a rectangular grid of typed or compact homogeneous values.",
  "fields" : [ {
    "name" : "rows",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellGridSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_BY_RANGE" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-range",
      "action" : {
        "type" : "SET_RANGE",
        "rows" : {
          "type" : "TYPED",
          "cells" : [ [ {
            "type" : "TEXT",
            "value" : {
              "type" : "INLINE",
              "text" : "Sample text"
            }
          } ] ]
        }
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1"
      }
    },
    "notes" : [ ]
  }
}
```

## CLEAR_RANGE

Clear all cell values and styles within the selected range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_BY_RANGE]]].

```json
{
  "id" : "CLEAR_RANGE",
  "summary" : "Clear all cell values and styles within the selected range.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_BY_RANGE" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-clear-range",
      "action" : {
        "type" : "CLEAR_RANGE"
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_ARRAY_FORMULA

Author one contiguous single-cell or multi-cell array-formula group.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"arrayFormulaSpecType"} |  |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_BY_RANGE]]].

```json
{
  "id" : "SET_ARRAY_FORMULA",
  "summary" : "Author one contiguous single-cell or multi-cell array-formula group.",
  "fields" : [ {
    "name" : "formula",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "arrayFormulaSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_BY_RANGE" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-array-formula",
      "action" : {
        "type" : "SET_ARRAY_FORMULA",
        "formula" : {
          "source" : {
            "type" : "INLINE",
            "text" : "Sample text"
          }
        }
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ ]
  }
}
```

## CLEAR_ARRAY_FORMULA

Remove the stored array-formula group targeted by any member cell.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_BY_ADDRESS]]].

```json
{
  "id" : "CLEAR_ARRAY_FORMULA",
  "summary" : "Remove the stored array-formula group targeted by any member cell.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_BY_ADDRESS" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-clear-array-formula",
      "action" : {
        "type" : "CLEAR_ARRAY_FORMULA"
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_HYPERLINK

Replace the hyperlink attached to a single cell.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| target | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"hyperlinkTargetTypes"} |  |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_BY_ADDRESS]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "SET_HYPERLINK",
  "summary" : "Replace the hyperlink attached to a single cell.",
  "fields" : [ {
    "name" : "target",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "hyperlinkTargetTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_BY_ADDRESS" ]
  }, {
    "family" : "TableCellSelector",
    "typeIds" : [ "TABLE_CELL_BY_COLUMN_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-hyperlink",
      "action" : {
        "type" : "SET_HYPERLINK",
        "target" : {
          "type" : "URL",
          "target" : "https://example.com"
        }
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; SET_HYPERLINK accepts CELL_BY_ADDRESS, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## CLEAR_HYPERLINK

Remove any hyperlink attached to a single existing cell.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_BY_ADDRESS]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "CLEAR_HYPERLINK",
  "summary" : "Remove any hyperlink attached to a single existing cell.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_BY_ADDRESS" ]
  }, {
    "family" : "TableCellSelector",
    "typeIds" : [ "TABLE_CELL_BY_COLUMN_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-clear-hyperlink",
      "action" : {
        "type" : "CLEAR_HYPERLINK"
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; CLEAR_HYPERLINK accepts CELL_BY_ADDRESS, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## SET_COMMENT

Replace the plain-text comment attached to a single cell.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| comment | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"commentSpecType"} |  |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_BY_ADDRESS]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "SET_COMMENT",
  "summary" : "Replace the plain-text comment attached to a single cell.",
  "fields" : [ {
    "name" : "comment",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "commentSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_BY_ADDRESS" ]
  }, {
    "family" : "TableCellSelector",
    "typeIds" : [ "TABLE_CELL_BY_COLUMN_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-comment",
      "action" : {
        "type" : "SET_COMMENT",
        "comment" : {
          "text" : {
            "type" : "INLINE",
            "text" : "Sample text"
          },
          "author" : "sample-author"
        }
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; SET_COMMENT accepts CELL_BY_ADDRESS, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## CLEAR_COMMENT

Remove any comment attached to a single existing cell.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_BY_ADDRESS]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "CLEAR_COMMENT",
  "summary" : "Remove any comment attached to a single existing cell.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_BY_ADDRESS" ]
  }, {
    "family" : "TableCellSelector",
    "typeIds" : [ "TABLE_CELL_BY_COLUMN_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-clear-comment",
      "action" : {
        "type" : "CLEAR_COMMENT"
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; CLEAR_COMMENT accepts CELL_BY_ADDRESS, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## APPLY_STYLE

Apply a style patch to every cell in the selected range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| style | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"cellStylePatchType"} |  |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_BY_RANGE]]].

```json
{
  "id" : "APPLY_STYLE",
  "summary" : "Apply a style patch to every cell in the selected range.",
  "fields" : [ {
    "name" : "style",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellStylePatchType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_BY_RANGE" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-apply-style",
      "action" : {
        "type" : "APPLY_STYLE",
        "style" : {
          "numberFormat" : "0.00"
        }
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ ]
  }
}
```

## APPEND_ROW

Append a new row of typed or compact homogeneous values after the last occupied row.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| values | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellRowSpecTypes"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "APPEND_ROW",
  "summary" : "Append a new row of typed or compact homogeneous values after the last occupied row.",
  "fields" : [ {
    "name" : "values",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellRowSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "STYLES", "HYPERLINKS", "COMMENTS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-append-row",
      "action" : {
        "type" : "APPEND_ROW",
        "values" : {
          "type" : "TYPED",
          "cells" : [ {
            "type" : "TEXT",
            "value" : {
              "type" : "INLINE",
              "text" : "Sample text"
            }
          } ]
        }
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_PICTURE

Create or replace one picture-backed drawing object on one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| picture | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"pictureSpecType"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_PICTURE",
  "summary" : "Create or replace one picture-backed drawing object on one sheet.",
  "fields" : [ {
    "name" : "picture",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "pictureSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-picture",
      "action" : {
        "type" : "SET_PICTURE",
        "picture" : {
          "name" : "Sheet1",
          "image" : {
            "format" : "EMF",
            "source" : {
              "type" : "INLINE_BASE64",
              "base64Data" : "AA=="
            }
          },
          "anchor" : {
            "type" : "TWO_CELL",
            "from" : {
              "columnIndex" : 0,
              "rowIndex" : 0,
              "dx" : 0,
              "dy" : 0
            },
            "to" : {
              "columnIndex" : 0,
              "rowIndex" : 0,
              "dx" : 0,
              "dy" : 0
            },
            "behavior" : "MOVE_AND_RESIZE"
          }
        }
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_SIGNATURE_LINE

Create or replace one signature-line drawing object on one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| signatureLine | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"signatureLineSpecType"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_SIGNATURE_LINE",
  "summary" : "Create or replace one signature-line drawing object on one sheet.",
  "fields" : [ {
    "name" : "signatureLine",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "signatureLineSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-signature-line",
      "action" : {
        "type" : "SET_SIGNATURE_LINE",
        "signatureLine" : {
          "name" : "Sheet1",
          "anchor" : {
            "type" : "TWO_CELL",
            "from" : {
              "columnIndex" : 0,
              "rowIndex" : 0,
              "dx" : 0,
              "dy" : 0
            },
            "to" : {
              "columnIndex" : 0,
              "rowIndex" : 0,
              "dx" : 0,
              "dy" : 0
            },
            "behavior" : "MOVE_AND_RESIZE"
          },
          "allowComments" : false,
          "caption" : "Sign here"
        }
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_CHART

Create or mutate one supported simple chart on one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| chart | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"chartSpecType"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_CHART",
  "summary" : "Create or mutate one supported simple chart on one sheet.",
  "fields" : [ {
    "name" : "chart",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "chartSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-chart",
      "action" : {
        "type" : "SET_CHART",
        "chart" : {
          "name" : "Sheet1",
          "anchor" : {
            "type" : "TWO_CELL",
            "from" : {
              "columnIndex" : 0,
              "rowIndex" : 0,
              "dx" : 0,
              "dy" : 0
            },
            "to" : {
              "columnIndex" : 0,
              "rowIndex" : 0,
              "dx" : 0,
              "dy" : 0
            },
            "behavior" : "MOVE_AND_RESIZE"
          },
          "title" : {
            "type" : "TEXT",
            "source" : {
              "type" : "INLINE",
              "text" : "Sample text"
            }
          },
          "legend" : {
            "type" : "HIDDEN"
          },
          "displayBlanksAs" : "GAP",
          "plotOnlyVisibleCells" : false,
          "plots" : [ {
            "type" : "AREA",
            "grouping" : "STANDARD",
            "axes" : [ {
              "kind" : "CATEGORY",
              "position" : "BOTTOM",
              "crosses" : "AUTO_ZERO",
              "visible" : true
            } ],
            "series" : [ {
              "title" : {
                "type" : "TEXT",
                "source" : {
                  "type" : "INLINE",
                  "text" : "Sample text"
                }
              },
              "categories" : {
                "type" : "REFERENCE",
                "formula" : "1"
              },
              "values" : {
                "type" : "REFERENCE",
                "formula" : "1"
              }
            } ]
          } ]
        }
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_SHAPE

Create or replace one simple-shape or connector drawing object on one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| shape | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"shapeSpecTypes"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_SHAPE",
  "summary" : "Create or replace one simple-shape or connector drawing object on one sheet.",
  "fields" : [ {
    "name" : "shape",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "shapeSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-shape",
      "action" : {
        "type" : "SET_SHAPE",
        "shape" : {
          "type" : "SIMPLE_SHAPE",
          "name" : "Sheet1",
          "anchor" : {
            "type" : "TWO_CELL",
            "from" : {
              "columnIndex" : 0,
              "rowIndex" : 0,
              "dx" : 0,
              "dy" : 0
            },
            "to" : {
              "columnIndex" : 0,
              "rowIndex" : 0,
              "dx" : 0,
              "dy" : 0
            },
            "behavior" : "MOVE_AND_RESIZE"
          },
          "presetGeometryToken" : "sample-presetGeometryToken"
        }
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_EMBEDDED_OBJECT

Create or replace one embedded-object drawing object on one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| embeddedObject | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"embeddedObjectSpecType"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "SET_EMBEDDED_OBJECT",
  "summary" : "Create or replace one embedded-object drawing object on one sheet.",
  "fields" : [ {
    "name" : "embeddedObject",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "embeddedObjectSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-embedded-object",
      "action" : {
        "type" : "SET_EMBEDDED_OBJECT",
        "embeddedObject" : {
          "name" : "Sheet1",
          "label" : "Sample label",
          "fileName" : "sample.bin",
          "command" : "sample-command",
          "payload" : {
            "type" : "INLINE_BASE64",
            "base64Data" : "AA=="
          },
          "previewImage" : {
            "format" : "EMF",
            "source" : {
              "type" : "INLINE_BASE64",
              "base64Data" : "AA=="
            }
          },
          "anchor" : {
            "type" : "TWO_CELL",
            "from" : {
              "columnIndex" : 0,
              "rowIndex" : 0,
              "dx" : 0,
              "dy" : 0
            },
            "to" : {
              "columnIndex" : 0,
              "rowIndex" : 0,
              "dx" : 0,
              "dy" : 0
            },
            "behavior" : "MOVE_AND_RESIZE"
          }
        }
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_DRAWING_OBJECT_ANCHOR

Move one existing drawing object by replacing its anchor authoritatively.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| anchor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingAnchorSpecTypes"} |  |

Targets: [TargetSelectorEntry[family=DrawingObjectSelector, typeIds=[DRAWING_OBJECT_BY_NAME]]].

```json
{
  "id" : "SET_DRAWING_OBJECT_ANCHOR",
  "summary" : "Move one existing drawing object by replacing its anchor authoritatively.",
  "fields" : [ {
    "name" : "anchor",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "drawingAnchorSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "DrawingObjectSelector",
    "typeIds" : [ "DRAWING_OBJECT_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-drawing-object-anchor",
      "action" : {
        "type" : "SET_DRAWING_OBJECT_ANCHOR",
        "anchor" : {
          "type" : "TWO_CELL",
          "from" : {
            "columnIndex" : 0,
            "rowIndex" : 0,
            "dx" : 0,
            "dy" : 0
          },
          "to" : {
            "columnIndex" : 0,
            "rowIndex" : 0,
            "dx" : 0,
            "dy" : 0
          },
          "behavior" : "MOVE_AND_RESIZE"
        }
      },
      "target" : {
        "type" : "DRAWING_OBJECT_BY_NAME",
        "sheetName" : "Sheet1",
        "objectName" : "sample-objectName"
      }
    },
    "notes" : [ ]
  }
}
```

## DELETE_DRAWING_OBJECT

Delete one existing drawing object by sheet-local name.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=DrawingObjectSelector, typeIds=[DRAWING_OBJECT_BY_NAME]]].

```json
{
  "id" : "DELETE_DRAWING_OBJECT",
  "summary" : "Delete one existing drawing object by sheet-local name.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "DrawingObjectSelector",
    "typeIds" : [ "DRAWING_OBJECT_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "DRAWINGS" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-delete-drawing-object",
      "action" : {
        "type" : "DELETE_DRAWING_OBJECT"
      },
      "target" : {
        "type" : "DRAWING_OBJECT_BY_NAME",
        "sheetName" : "Sheet1",
        "objectName" : "sample-objectName"
      }
    },
    "notes" : [ ]
  }
}
```

## IMPORT_CUSTOM_XML_MAPPING

Import one XML document into one existing workbook custom-XML mapping.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| mapping | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"customXmlImportSpecType"} |  |

Targets: [TargetSelectorEntry[family=WorkbookSelector, typeIds=[WORKBOOK_CURRENT]]].

```json
{
  "id" : "IMPORT_CUSTOM_XML_MAPPING",
  "summary" : "Import one XML document into one existing workbook custom-XML mapping.",
  "fields" : [ {
    "name" : "mapping",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "customXmlImportSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "WorkbookSelector",
    "typeIds" : [ "WORKBOOK_CURRENT" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-import-custom-xml-mapping",
      "action" : {
        "type" : "IMPORT_CUSTOM_XML_MAPPING",
        "mapping" : {
          "locator" : {
            "name" : "Mapping1"
          },
          "xml" : {
            "type" : "INLINE",
            "text" : "Sample text"
          }
        }
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_PIVOT_TABLE

Create or replace one workbook-global pivot-table definition.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pivotTable | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"pivotTableSpecType"} |  |

Targets: [TargetSelectorEntry[family=PivotTableSelector, typeIds=[PIVOT_TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "SET_PIVOT_TABLE",
  "summary" : "Create or replace one workbook-global pivot-table definition.",
  "fields" : [ {
    "name" : "pivotTable",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "pivotTableSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "PivotTableSelector",
    "typeIds" : [ "PIVOT_TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-pivot-table",
      "action" : {
        "type" : "SET_PIVOT_TABLE",
        "pivotTable" : {
          "name" : "Sheet1",
          "sheetName" : "Sheet1",
          "source" : {
            "type" : "RANGE",
            "sheetName" : "Sheet1",
            "range" : "A1:B2"
          },
          "anchor" : {
            "topLeftAddress" : "A1"
          },
          "dataFields" : [ {
            "sourceColumnName" : "sample-sourceColumnName",
            "function" : "SUM"
          } ]
        }
      },
      "target" : {
        "type" : "PIVOT_TABLE_BY_NAME_ON_SHEET",
        "name" : "Sheet1",
        "sheetName" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_DATA_VALIDATION

Create or replace one data-validation rule over the requested sheet range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| validation | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"dataValidationSpecType"} |  |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_BY_RANGE]]].

```json
{
  "id" : "SET_DATA_VALIDATION",
  "summary" : "Create or replace one data-validation rule over the requested sheet range.",
  "fields" : [ {
    "name" : "validation",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "dataValidationSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_BY_RANGE" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-data-validation",
      "action" : {
        "type" : "SET_DATA_VALIDATION",
        "validation" : {
          "rule" : {
            "type" : "EXPLICIT_LIST",
            "values" : [ "sample-values" ]
          }
        }
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ ]
  }
}
```

## CLEAR_DATA_VALIDATIONS

Remove data-validation structures on the sheet that match the range selection.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_ALL_ON_SHEET, RANGE_BY_RANGE, RANGE_BY_RANGES, RANGE_RECTANGULAR_WINDOW]]].

```json
{
  "id" : "CLEAR_DATA_VALIDATIONS",
  "summary" : "Remove data-validation structures on the sheet that match the range selection.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_ALL_ON_SHEET", "RANGE_BY_RANGE", "RANGE_BY_RANGES", "RANGE_RECTANGULAR_WINDOW" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-clear-data-validations",
      "action" : {
        "type" : "CLEAR_DATA_VALIDATIONS"
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; CLEAR_DATA_VALIDATIONS accepts RANGE_ALL_ON_SHEET, RANGE_BY_RANGE, RANGE_BY_RANGES, RANGE_RECTANGULAR_WINDOW." ]
  }
}
```

## SET_CONDITIONAL_FORMATTING

Create or replace one logical conditional-formatting block over explicit sheet ranges.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| conditionalFormatting | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"conditionalFormattingSpecType"} |  |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_BY_RANGE, RANGE_BY_RANGES]]].

```json
{
  "id" : "SET_CONDITIONAL_FORMATTING",
  "summary" : "Create or replace one logical conditional-formatting block over explicit sheet ranges.",
  "fields" : [ {
    "name" : "conditionalFormatting",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "conditionalFormattingSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_BY_RANGE", "RANGE_BY_RANGES" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-conditional-formatting",
      "action" : {
        "type" : "SET_CONDITIONAL_FORMATTING",
        "conditionalFormatting" : {
          "rules" : [ {
            "type" : "FORMULA_RULE",
            "formula" : "1",
            "stopIfTrue" : true
          } ]
        }
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; SET_CONDITIONAL_FORMATTING accepts RANGE_BY_RANGE, RANGE_BY_RANGES." ]
  }
}
```

## CLEAR_CONDITIONAL_FORMATTING

Remove conditional-formatting blocks on the sheet that match the range selection.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_ALL_ON_SHEET, RANGE_BY_RANGE, RANGE_BY_RANGES, RANGE_RECTANGULAR_WINDOW]]].

```json
{
  "id" : "CLEAR_CONDITIONAL_FORMATTING",
  "summary" : "Remove conditional-formatting blocks on the sheet that match the range selection.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_ALL_ON_SHEET", "RANGE_BY_RANGE", "RANGE_BY_RANGES", "RANGE_RECTANGULAR_WINDOW" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-clear-conditional-formatting",
      "action" : {
        "type" : "CLEAR_CONDITIONAL_FORMATTING"
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; CLEAR_CONDITIONAL_FORMATTING accepts RANGE_ALL_ON_SHEET, RANGE_BY_RANGE, RANGE_BY_RANGES, RANGE_RECTANGULAR_WINDOW." ]
  }
}
```

## SET_AUTOFILTER

Create or replace one sheet-level autofilter range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| criteria | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"autofilterFilterColumnSpecType"}} |  |
| sortState | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"autofilterSortStateSpecType"} |  |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_BY_RANGE]]].

```json
{
  "id" : "SET_AUTOFILTER",
  "summary" : "Create or replace one sheet-level autofilter range.",
  "fields" : [ {
    "name" : "criteria",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "autofilterFilterColumnSpecType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "sortState",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "autofilterSortStateSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_BY_RANGE" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-autofilter",
      "action" : {
        "type" : "SET_AUTOFILTER",
        "criteria" : [ {
          "columnId" : 1,
          "criterion" : {
            "type" : "VALUES",
            "values" : [ "sample-values" ],
            "includeBlank" : false
          }
        } ]
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ ]
  }
}
```

## CLEAR_AUTOFILTER

Clear the sheet-level autofilter range on one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "CLEAR_AUTOFILTER",
  "summary" : "Clear the sheet-level autofilter range on one sheet.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-clear-autofilter",
      "action" : {
        "type" : "CLEAR_AUTOFILTER"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_TABLE

Create or replace one workbook-global table definition.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| table | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"tableSpecType"} |  |

Targets: [TargetSelectorEntry[family=TableSelector, typeIds=[TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "SET_TABLE",
  "summary" : "Create or replace one workbook-global table definition.",
  "fields" : [ {
    "name" : "table",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "tableSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "TableSelector",
    "typeIds" : [ "TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-table",
      "action" : {
        "type" : "SET_TABLE",
        "table" : {
          "name" : "Sheet1",
          "sheetName" : "Sheet1",
          "range" : "A1:B2",
          "style" : {
            "type" : "NONE"
          }
        }
      },
      "target" : {
        "type" : "TABLE_BY_NAME_ON_SHEET",
        "name" : "Sheet1",
        "sheetName" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## DELETE_TABLE

Delete one existing table by workbook-global name and expected sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=TableSelector, typeIds=[TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "DELETE_TABLE",
  "summary" : "Delete one existing table by workbook-global name and expected sheet.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "TableSelector",
    "typeIds" : [ "TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-delete-table",
      "action" : {
        "type" : "DELETE_TABLE"
      },
      "target" : {
        "type" : "TABLE_BY_NAME_ON_SHEET",
        "name" : "Sheet1",
        "sheetName" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## DELETE_PIVOT_TABLE

Delete one existing pivot table by workbook-global name and expected sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=PivotTableSelector, typeIds=[PIVOT_TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "DELETE_PIVOT_TABLE",
  "summary" : "Delete one existing pivot table by workbook-global name and expected sheet.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "PivotTableSelector",
    "typeIds" : [ "PIVOT_TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-delete-pivot-table",
      "action" : {
        "type" : "DELETE_PIVOT_TABLE"
      },
      "target" : {
        "type" : "PIVOT_TABLE_BY_NAME_ON_SHEET",
        "name" : "Sheet1",
        "sheetName" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## SET_NAMED_RANGE

Create or replace one typed named range in workbook or sheet scope.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"LENGTH_RANGE","minimum":1,"maximum":255},{"type":"STRING_PATTERN","pattern":"^[\\p{L}_\\\\][\\p{L}\\p{N}_.\\\\]*$"}]} |  |
| scope | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"namedRangeScopeTypes"} |  |
| target | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"namedRangeTargetTypes"} |  |

Targets: [TargetSelectorEntry[family=NamedRangeSelector, typeIds=[NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE]]].

```json
{
  "id" : "SET_NAMED_RANGE",
  "summary" : "Create or replace one typed named range in workbook or sheet scope.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "LENGTH_RANGE",
        "minimum" : 1,
        "maximum" : 255
      }, {
        "type" : "STRING_PATTERN",
        "pattern" : "^[\\p{L}_\\\\][\\p{L}\\p{N}_.\\\\]*$"
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "scope",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "namedRangeScopeTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "target",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "namedRangeTargetTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "NamedRangeSelector",
    "typeIds" : [ "NAMED_RANGE_WORKBOOK_SCOPE", "NAMED_RANGE_SHEET_SCOPE" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-set-named-range",
      "action" : {
        "type" : "SET_NAMED_RANGE",
        "name" : "RevenueRange",
        "scope" : {
          "type" : "WORKBOOK"
        },
        "target" : {
          "type" : "RANGE",
          "sheetName" : "Sheet1",
          "range" : "A1:B2"
        }
      },
      "target" : {
        "type" : "NAMED_RANGE_WORKBOOK_SCOPE",
        "name" : "RevenueRange"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; SET_NAMED_RANGE accepts NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE." ]
  }
}
```

## DELETE_NAMED_RANGE

Delete one existing named range from workbook or sheet scope.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=NamedRangeSelector, typeIds=[NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE]]].

```json
{
  "id" : "DELETE_NAMED_RANGE",
  "summary" : "Delete one existing named range from workbook or sheet scope.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "NamedRangeSelector",
    "typeIds" : [ "NAMED_RANGE_WORKBOOK_SCOPE", "NAMED_RANGE_SHEET_SCOPE" ]
  } ],
  "effects" : [ "READ_WORKBOOK", "MUTATE_WORKBOOK", "CELLS", "FORMULAS", "TABLES", "NAMED_RANGES", "DATA_VALIDATIONS", "CONDITIONAL_FORMATTING", "AUTOFILTERS", "PIVOT_TABLES", "CUSTOM_XML" ],
  "effectFootprint" : "CONSERVATIVE_WORKBOOK_WIDE",
  "stepTemplate" : {
    "stepKind" : "MUTATION",
    "bodyField" : "action",
    "template" : {
      "stepId" : "step-delete-named-range",
      "action" : {
        "type" : "DELETE_NAMED_RANGE"
      },
      "target" : {
        "type" : "NAMED_RANGE_WORKBOOK_SCOPE",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; DELETE_NAMED_RANGE accepts NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE." ]
  }
}
```
