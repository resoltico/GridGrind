<!-- Generated from core metadata; do not edit. -->
# assertionTypes

## EXPECT_SHEET_PRESENT

Require the selected sheet selector to resolve to one or more sheets.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES]]].

```json
{
  "id" : "EXPECT_SHEET_PRESENT",
  "summary" : "Require the selected sheet selector to resolve to one or more sheets.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_ALL", "SHEET_BY_NAME", "SHEET_BY_NAMES" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-sheet-present",
      "assertion" : {
        "type" : "EXPECT_SHEET_PRESENT"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_SHEET_PRESENT accepts SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES." ]
  }
}
```

## EXPECT_SHEET_ABSENT

Require the selected sheet selector to resolve to no sheets.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES]]].

```json
{
  "id" : "EXPECT_SHEET_ABSENT",
  "summary" : "Require the selected sheet selector to resolve to no sheets.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_ALL", "SHEET_BY_NAME", "SHEET_BY_NAMES" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-sheet-absent",
      "assertion" : {
        "type" : "EXPECT_SHEET_ABSENT"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_SHEET_ABSENT accepts SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES." ]
  }
}
```

## EXPECT_NAMED_RANGE_PRESENT

Require the selected named-range selector to resolve to one or more named ranges.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=NamedRangeSelector, typeIds=[NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE]]].

```json
{
  "id" : "EXPECT_NAMED_RANGE_PRESENT",
  "summary" : "Require the selected named-range selector to resolve to one or more named ranges.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "NamedRangeSelector",
    "typeIds" : [ "NAMED_RANGE_ALL", "NAMED_RANGE_ANY_OF", "NAMED_RANGE_BY_NAMES", "NAMED_RANGE_BY_NAME", "NAMED_RANGE_WORKBOOK_SCOPE", "NAMED_RANGE_SHEET_SCOPE" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-named-range-present",
      "assertion" : {
        "type" : "EXPECT_NAMED_RANGE_PRESENT"
      },
      "target" : {
        "type" : "NAMED_RANGE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_NAMED_RANGE_PRESENT accepts NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE." ]
  }
}
```

## EXPECT_NAMED_RANGE_ABSENT

Require the selected named-range selector to resolve to no named ranges.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=NamedRangeSelector, typeIds=[NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE]]].

```json
{
  "id" : "EXPECT_NAMED_RANGE_ABSENT",
  "summary" : "Require the selected named-range selector to resolve to no named ranges.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "NamedRangeSelector",
    "typeIds" : [ "NAMED_RANGE_ALL", "NAMED_RANGE_ANY_OF", "NAMED_RANGE_BY_NAMES", "NAMED_RANGE_BY_NAME", "NAMED_RANGE_WORKBOOK_SCOPE", "NAMED_RANGE_SHEET_SCOPE" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-named-range-absent",
      "assertion" : {
        "type" : "EXPECT_NAMED_RANGE_ABSENT"
      },
      "target" : {
        "type" : "NAMED_RANGE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_NAMED_RANGE_ABSENT accepts NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE." ]
  }
}
```

## EXPECT_TABLE_PRESENT

Require the selected table selector to resolve to one or more tables.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=TableSelector, typeIds=[TABLE_ALL, TABLE_BY_NAME, TABLE_BY_NAMES, TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "EXPECT_TABLE_PRESENT",
  "summary" : "Require the selected table selector to resolve to one or more tables.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "TableSelector",
    "typeIds" : [ "TABLE_ALL", "TABLE_BY_NAME", "TABLE_BY_NAMES", "TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-table-present",
      "assertion" : {
        "type" : "EXPECT_TABLE_PRESENT"
      },
      "target" : {
        "type" : "TABLE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_TABLE_PRESENT accepts TABLE_ALL, TABLE_BY_NAME, TABLE_BY_NAMES, TABLE_BY_NAME_ON_SHEET." ]
  }
}
```

## EXPECT_TABLE_ABSENT

Require the selected table selector to resolve to no tables.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=TableSelector, typeIds=[TABLE_ALL, TABLE_BY_NAME, TABLE_BY_NAMES, TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "EXPECT_TABLE_ABSENT",
  "summary" : "Require the selected table selector to resolve to no tables.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "TableSelector",
    "typeIds" : [ "TABLE_ALL", "TABLE_BY_NAME", "TABLE_BY_NAMES", "TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-table-absent",
      "assertion" : {
        "type" : "EXPECT_TABLE_ABSENT"
      },
      "target" : {
        "type" : "TABLE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_TABLE_ABSENT accepts TABLE_ALL, TABLE_BY_NAME, TABLE_BY_NAMES, TABLE_BY_NAME_ON_SHEET." ]
  }
}
```

## EXPECT_PIVOT_TABLE_PRESENT

Require the selected pivot-table selector to resolve to one or more pivot tables.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=PivotTableSelector, typeIds=[PIVOT_TABLE_ALL, PIVOT_TABLE_BY_NAME, PIVOT_TABLE_BY_NAMES, PIVOT_TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "EXPECT_PIVOT_TABLE_PRESENT",
  "summary" : "Require the selected pivot-table selector to resolve to one or more pivot tables.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "PivotTableSelector",
    "typeIds" : [ "PIVOT_TABLE_ALL", "PIVOT_TABLE_BY_NAME", "PIVOT_TABLE_BY_NAMES", "PIVOT_TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-pivot-table-present",
      "assertion" : {
        "type" : "EXPECT_PIVOT_TABLE_PRESENT"
      },
      "target" : {
        "type" : "PIVOT_TABLE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_PIVOT_TABLE_PRESENT accepts PIVOT_TABLE_ALL, PIVOT_TABLE_BY_NAME, PIVOT_TABLE_BY_NAMES, PIVOT_TABLE_BY_NAME_ON_SHEET." ]
  }
}
```

## EXPECT_PIVOT_TABLE_ABSENT

Require the selected pivot-table selector to resolve to no pivot tables.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=PivotTableSelector, typeIds=[PIVOT_TABLE_ALL, PIVOT_TABLE_BY_NAME, PIVOT_TABLE_BY_NAMES, PIVOT_TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "EXPECT_PIVOT_TABLE_ABSENT",
  "summary" : "Require the selected pivot-table selector to resolve to no pivot tables.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "PivotTableSelector",
    "typeIds" : [ "PIVOT_TABLE_ALL", "PIVOT_TABLE_BY_NAME", "PIVOT_TABLE_BY_NAMES", "PIVOT_TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-pivot-table-absent",
      "assertion" : {
        "type" : "EXPECT_PIVOT_TABLE_ABSENT"
      },
      "target" : {
        "type" : "PIVOT_TABLE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_PIVOT_TABLE_ABSENT accepts PIVOT_TABLE_ALL, PIVOT_TABLE_BY_NAME, PIVOT_TABLE_BY_NAMES, PIVOT_TABLE_BY_NAME_ON_SHEET." ]
  }
}
```

## EXPECT_CHART_PRESENT

Require the selected chart selector to resolve to one or more charts.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=ChartSelector, typeIds=[CHART_ALL_ON_SHEET, CHART_BY_NAME]]].

```json
{
  "id" : "EXPECT_CHART_PRESENT",
  "summary" : "Require the selected chart selector to resolve to one or more charts.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "ChartSelector",
    "typeIds" : [ "CHART_ALL_ON_SHEET", "CHART_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-chart-present",
      "assertion" : {
        "type" : "EXPECT_CHART_PRESENT"
      },
      "target" : {
        "type" : "CHART_BY_NAME",
        "sheetName" : "Sheet1",
        "chartName" : "sample-chartName"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_CHART_PRESENT accepts CHART_ALL_ON_SHEET, CHART_BY_NAME." ]
  }
}
```

## EXPECT_CHART_ABSENT

Require the selected chart selector to resolve to no charts.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=ChartSelector, typeIds=[CHART_ALL_ON_SHEET, CHART_BY_NAME]]].

```json
{
  "id" : "EXPECT_CHART_ABSENT",
  "summary" : "Require the selected chart selector to resolve to no charts.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "ChartSelector",
    "typeIds" : [ "CHART_ALL_ON_SHEET", "CHART_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-chart-absent",
      "assertion" : {
        "type" : "EXPECT_CHART_ABSENT"
      },
      "target" : {
        "type" : "CHART_BY_NAME",
        "sheetName" : "Sheet1",
        "chartName" : "sample-chartName"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_CHART_ABSENT accepts CHART_ALL_ON_SHEET, CHART_BY_NAME." ]
  }
}
```

## EXPECT_CELL_VALUE

Require every selected cell to have the exact effective value.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| expectedValue | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellValueTypes"} |  |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_BY_ADDRESS, CELL_BY_ADDRESSES]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "EXPECT_CELL_VALUE",
  "summary" : "Require every selected cell to have the exact effective value.",
  "fields" : [ {
    "name" : "expectedValue",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellValueTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_BY_ADDRESS", "CELL_BY_ADDRESSES" ]
  }, {
    "family" : "TableCellSelector",
    "typeIds" : [ "TABLE_CELL_BY_COLUMN_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-cell-value",
      "assertion" : {
        "type" : "EXPECT_CELL_VALUE",
        "expectedValue" : {
          "type" : "TEXT",
          "value" : "sample-value"
        }
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_CELL_VALUE accepts CELL_BY_ADDRESS, CELL_BY_ADDRESSES, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## EXPECT_DISPLAY_VALUE

Require every selected cell to have the exact formatted display string.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| displayValue | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_BY_ADDRESS, CELL_BY_ADDRESSES]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "EXPECT_DISPLAY_VALUE",
  "summary" : "Require every selected cell to have the exact formatted display string.",
  "fields" : [ {
    "name" : "displayValue",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_BY_ADDRESS", "CELL_BY_ADDRESSES" ]
  }, {
    "family" : "TableCellSelector",
    "typeIds" : [ "TABLE_CELL_BY_COLUMN_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-display-value",
      "assertion" : {
        "type" : "EXPECT_DISPLAY_VALUE",
        "displayValue" : "sample-displayValue"
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_DISPLAY_VALUE accepts CELL_BY_ADDRESS, CELL_BY_ADDRESSES, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## EXPECT_FORMULA_TEXT

Require every selected cell to store the exact formula text.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_BY_ADDRESS, CELL_BY_ADDRESSES]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "EXPECT_FORMULA_TEXT",
  "summary" : "Require every selected cell to store the exact formula text.",
  "fields" : [ {
    "name" : "formula",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_BY_ADDRESS", "CELL_BY_ADDRESSES" ]
  }, {
    "family" : "TableCellSelector",
    "typeIds" : [ "TABLE_CELL_BY_COLUMN_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-formula-text",
      "assertion" : {
        "type" : "EXPECT_FORMULA_TEXT",
        "formula" : "1"
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_FORMULA_TEXT accepts CELL_BY_ADDRESS, CELL_BY_ADDRESSES, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## EXPECT_CELL_STYLE

Require every selected cell to have the exact style snapshot.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| style | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"cellStyleReportType"} |  |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_BY_ADDRESS, CELL_BY_ADDRESSES]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "EXPECT_CELL_STYLE",
  "summary" : "Require every selected cell to have the exact style snapshot.",
  "fields" : [ {
    "name" : "style",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellStyleReportType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_BY_ADDRESS", "CELL_BY_ADDRESSES" ]
  }, {
    "family" : "TableCellSelector",
    "typeIds" : [ "TABLE_CELL_BY_COLUMN_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-cell-style",
      "assertion" : {
        "type" : "EXPECT_CELL_STYLE",
        "style" : {
          "numberFormat" : "sample-numberFormat",
          "alignment" : {
            "wrapText" : false,
            "horizontalAlignment" : "GENERAL",
            "verticalAlignment" : "TOP",
            "textRotation" : 1,
            "indentation" : 1
          },
          "font" : {
            "bold" : false,
            "italic" : false,
            "fontName" : "sample-fontName",
            "fontHeight" : {
              "twips" : 20,
              "points" : 1
            },
            "underline" : false,
            "strikeout" : false
          },
          "fill" : {
            "type" : "PATTERN_ONLY",
            "pattern" : "NONE"
          },
          "border" : {
            "top" : {
              "type" : "NONE"
            },
            "right" : {
              "type" : "NONE"
            },
            "bottom" : {
              "type" : "NONE"
            },
            "left" : {
              "type" : "NONE"
            }
          },
          "protection" : {
            "locked" : true,
            "hiddenFormula" : false
          }
        }
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_CELL_STYLE accepts CELL_BY_ADDRESS, CELL_BY_ADDRESSES, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## EXPECT_WORKBOOK_PROTECTION

Require the workbook protection report to match exactly.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| protection | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"workbookProtectionReportType"} |  |

Targets: [TargetSelectorEntry[family=WorkbookSelector, typeIds=[WORKBOOK_CURRENT]]].

```json
{
  "id" : "EXPECT_WORKBOOK_PROTECTION",
  "summary" : "Require the workbook protection report to match exactly.",
  "fields" : [ {
    "name" : "protection",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "workbookProtectionReportType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "WorkbookSelector",
    "typeIds" : [ "WORKBOOK_CURRENT" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-workbook-protection",
      "assertion" : {
        "type" : "EXPECT_WORKBOOK_PROTECTION",
        "protection" : {
          "structureLocked" : false,
          "windowsLocked" : false,
          "revisionsLocked" : false,
          "workbookPasswordHashPresent" : false,
          "revisionsPasswordHashPresent" : false
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

## EXPECT_SHEET_STRUCTURE

Require the selected sheet summary report to match exactly.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheet | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"sheetSummaryReportType"} |  |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "EXPECT_SHEET_STRUCTURE",
  "summary" : "Require the selected sheet summary report to match exactly.",
  "fields" : [ {
    "name" : "sheet",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "sheetSummaryReportType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-sheet-structure",
      "assertion" : {
        "type" : "EXPECT_SHEET_STRUCTURE",
        "sheet" : {
          "sheetName" : "Sheet1",
          "visibility" : "VISIBLE",
          "protection" : {
            "type" : "UNPROTECTED"
          },
          "physicalRowCount" : 1,
          "lastRowIndex" : 0,
          "lastColumnIndex" : 0
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

## EXPECT_NAMED_RANGE_FACTS

Require the selected named-range reports to match exactly and in order.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| namedRanges | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"namedRangeReportTypes"}} |  |

Targets: [TargetSelectorEntry[family=NamedRangeSelector, typeIds=[NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE]]].

```json
{
  "id" : "EXPECT_NAMED_RANGE_FACTS",
  "summary" : "Require the selected named-range reports to match exactly and in order.",
  "fields" : [ {
    "name" : "namedRanges",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "namedRangeReportTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "NamedRangeSelector",
    "typeIds" : [ "NAMED_RANGE_ALL", "NAMED_RANGE_ANY_OF", "NAMED_RANGE_BY_NAMES", "NAMED_RANGE_BY_NAME", "NAMED_RANGE_WORKBOOK_SCOPE", "NAMED_RANGE_SHEET_SCOPE" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-named-range-facts",
      "assertion" : {
        "type" : "EXPECT_NAMED_RANGE_FACTS",
        "namedRanges" : [ {
          "type" : "RANGE",
          "name" : "Sheet1",
          "scope" : {
            "type" : "WORKBOOK"
          },
          "refersToFormula" : "sample-refersToFormula",
          "target" : {
            "type" : "FORMULA",
            "formula" : "1"
          }
        } ]
      },
      "target" : {
        "type" : "NAMED_RANGE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_NAMED_RANGE_FACTS accepts NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE." ]
  }
}
```

## EXPECT_TABLE_FACTS

Require the selected table reports to match exactly and in order.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| tables | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"tableEntryReportType"}} |  |

Targets: [TargetSelectorEntry[family=TableSelector, typeIds=[TABLE_ALL, TABLE_BY_NAME, TABLE_BY_NAMES, TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "EXPECT_TABLE_FACTS",
  "summary" : "Require the selected table reports to match exactly and in order.",
  "fields" : [ {
    "name" : "tables",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "tableEntryReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "TableSelector",
    "typeIds" : [ "TABLE_ALL", "TABLE_BY_NAME", "TABLE_BY_NAMES", "TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-table-facts",
      "assertion" : {
        "type" : "EXPECT_TABLE_FACTS",
        "tables" : [ {
          "name" : "Sheet1",
          "sheetName" : "Sheet1",
          "range" : "A1:B2",
          "structure" : {
            "headerRowCount" : 1,
            "totalsRowCount" : 1,
            "columnNames" : [ "sample-columnNames" ],
            "columns" : [ {
              "id" : 1,
              "name" : "Sheet1"
            } ]
          },
          "style" : {
            "type" : "NONE"
          }
        } ]
      },
      "target" : {
        "type" : "TABLE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_TABLE_FACTS accepts TABLE_ALL, TABLE_BY_NAME, TABLE_BY_NAMES, TABLE_BY_NAME_ON_SHEET." ]
  }
}
```

## EXPECT_PIVOT_TABLE_FACTS

Require the selected pivot-table reports to match exactly and in order.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pivotTables | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"pivotTableReportTypes"}} |  |

Targets: [TargetSelectorEntry[family=PivotTableSelector, typeIds=[PIVOT_TABLE_ALL, PIVOT_TABLE_BY_NAME, PIVOT_TABLE_BY_NAMES, PIVOT_TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "EXPECT_PIVOT_TABLE_FACTS",
  "summary" : "Require the selected pivot-table reports to match exactly and in order.",
  "fields" : [ {
    "name" : "pivotTables",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "pivotTableReportTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "PivotTableSelector",
    "typeIds" : [ "PIVOT_TABLE_ALL", "PIVOT_TABLE_BY_NAME", "PIVOT_TABLE_BY_NAMES", "PIVOT_TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-pivot-table-facts",
      "assertion" : {
        "type" : "EXPECT_PIVOT_TABLE_FACTS",
        "pivotTables" : [ {
          "type" : "SUPPORTED",
          "name" : "Sheet1",
          "sheetName" : "Sheet1",
          "anchor" : {
            "topLeftAddress" : "A1",
            "locationRange" : "sample-locationRange"
          },
          "source" : {
            "type" : "RANGE",
            "sheetName" : "Sheet1",
            "range" : "A1:B2"
          },
          "rowLabels" : [ {
            "sourceColumnIndex" : 1,
            "sourceColumnName" : "sample-sourceColumnName"
          } ],
          "columnLabels" : [ {
            "sourceColumnIndex" : 1,
            "sourceColumnName" : "sample-sourceColumnName"
          } ],
          "reportFilters" : [ {
            "sourceColumnIndex" : 1,
            "sourceColumnName" : "sample-sourceColumnName"
          } ],
          "dataFields" : [ {
            "sourceColumnIndex" : 1,
            "sourceColumnName" : "sample-sourceColumnName",
            "function" : "SUM",
            "displayName" : "Sample name"
          } ],
          "valuesAxisOnColumns" : false
        } ]
      },
      "target" : {
        "type" : "PIVOT_TABLE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_PIVOT_TABLE_FACTS accepts PIVOT_TABLE_ALL, PIVOT_TABLE_BY_NAME, PIVOT_TABLE_BY_NAMES, PIVOT_TABLE_BY_NAME_ON_SHEET." ]
  }
}
```

## EXPECT_CHART_FACTS

Require the selected chart reports to match exactly and in order.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| charts | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"chartReportType"}} |  |

Targets: [TargetSelectorEntry[family=ChartSelector, typeIds=[CHART_ALL_ON_SHEET, CHART_BY_NAME]]].

```json
{
  "id" : "EXPECT_CHART_FACTS",
  "summary" : "Require the selected chart reports to match exactly and in order.",
  "fields" : [ {
    "name" : "charts",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "chartReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "ChartSelector",
    "typeIds" : [ "CHART_ALL_ON_SHEET", "CHART_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-chart-facts",
      "assertion" : {
        "type" : "EXPECT_CHART_FACTS",
        "charts" : [ {
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
            "text" : "Sample text"
          },
          "legend" : {
            "type" : "HIDDEN"
          },
          "displayBlanksAs" : "GAP",
          "plotOnlyVisibleCells" : false,
          "plots" : [ {
            "type" : "AREA",
            "varyColors" : false,
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
                "text" : "Sample text"
              },
              "categories" : {
                "type" : "STRING_REFERENCE",
                "formula" : "1",
                "cachedValues" : [ "sample-cachedValues" ]
              },
              "values" : {
                "type" : "STRING_REFERENCE",
                "formula" : "1",
                "cachedValues" : [ "sample-cachedValues" ]
              }
            } ]
          } ]
        } ]
      },
      "target" : {
        "type" : "CHART_BY_NAME",
        "sheetName" : "Sheet1",
        "chartName" : "sample-chartName"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; EXPECT_CHART_FACTS accepts CHART_ALL_ON_SHEET, CHART_BY_NAME." ]
  }
}
```

## EXPECT_ANALYSIS_MAX_SEVERITY

Run one analysis query and require its highest severity to be no higher than maximumSeverity.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| query | REQUIRED | {"type":"TOP_LEVEL_TYPE_SET","typeSet":"inspectionQueryTypes"} |  |
| maximumSeverity | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "EXPECT_ANALYSIS_MAX_SEVERITY",
  "summary" : "Run one analysis query and require its highest severity to be no higher than maximumSeverity.",
  "fields" : [ {
    "name" : "query",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "TOP_LEVEL_TYPE_SET",
      "typeSet" : "inspectionQueryTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "maximumSeverity",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "INFO", "WARNING", "ERROR" ]
  } ],
  "targetSelectors" : [ ],
  "targetSelectorRule" : "Matches the nested analysis query's target selectors.",
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-analysis-max-severity",
      "assertion" : {
        "type" : "EXPECT_ANALYSIS_MAX_SEVERITY",
        "query" : {
          "type" : "ANALYZE_WORKBOOK_FINDINGS"
        },
        "maximumSeverity" : "INFO"
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ "Target rule: Matches the nested analysis query's target selectors." ]
  }
}
```

## EXPECT_ANALYSIS_FINDING_PRESENT

Run one analysis query and require at least one matching finding.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| query | REQUIRED | {"type":"TOP_LEVEL_TYPE_SET","typeSet":"inspectionQueryTypes"} |  |
| code | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| severity | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| messageContains | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "EXPECT_ANALYSIS_FINDING_PRESENT",
  "summary" : "Run one analysis query and require at least one matching finding.",
  "fields" : [ {
    "name" : "query",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "TOP_LEVEL_TYPE_SET",
      "typeSet" : "inspectionQueryTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "code",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "FORMULA_ERROR_RESULT", "FORMULA_EVALUATION_FAILURE", "FORMULA_EXTERNAL_REFERENCE", "FORMULA_MISSING_EXTERNAL_WORKBOOK", "FORMULA_USES_CACHED_EXTERNAL_VALUE", "FORMULA_UNREGISTERED_USER_DEFINED_FUNCTION", "FORMULA_VOLATILE_FUNCTION", "DATA_VALIDATION_EMPTY_EXPLICIT_LIST", "DATA_VALIDATION_MALFORMED_RULE", "DATA_VALIDATION_UNSUPPORTED_RULE", "DATA_VALIDATION_BROKEN_FORMULA", "DATA_VALIDATION_OVERLAPPING_RULES", "CONDITIONAL_FORMATTING_BROKEN_FORMULA", "CONDITIONAL_FORMATTING_UNSUPPORTED_RULE", "CONDITIONAL_FORMATTING_PRIORITY_COLLISION", "CONDITIONAL_FORMATTING_EMPTY_RANGE", "AUTOFILTER_INVALID_RANGE", "AUTOFILTER_MISSING_HEADER_ROW", "AUTOFILTER_TABLE_MISMATCH", "TABLE_OVERLAPPING_RANGE", "TABLE_BLANK_HEADER", "TABLE_DUPLICATE_HEADER", "TABLE_BROKEN_REFERENCE", "TABLE_STYLE_MISMATCH", "PIVOT_TABLE_MISSING_NAME", "PIVOT_TABLE_DUPLICATE_NAME", "PIVOT_TABLE_MISSING_CACHE_DEFINITION", "PIVOT_TABLE_MISSING_WORKBOOK_CACHE", "PIVOT_TABLE_MISSING_CACHE_RECORDS", "PIVOT_TABLE_BROKEN_SOURCE", "PIVOT_TABLE_UNSUPPORTED_DETAIL", "HYPERLINK_MALFORMED_TARGET", "HYPERLINK_MISSING_FILE_TARGET", "HYPERLINK_MISSING_DOCUMENT_SHEET", "HYPERLINK_UNRESOLVED_FILE_TARGET", "HYPERLINK_INVALID_DOCUMENT_TARGET", "NAMED_RANGE_BROKEN_REFERENCE", "NAMED_RANGE_UNPARSEABLE_FORMULA", "NAMED_RANGE_UNRESOLVED_TARGET", "NAMED_RANGE_SCOPE_SHADOWING" ]
  }, {
    "name" : "severity",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "INFO", "WARNING", "ERROR" ]
  }, {
    "name" : "messageContains",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ],
  "targetSelectorRule" : "Matches the nested analysis query's target selectors.",
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-analysis-finding-present",
      "assertion" : {
        "type" : "EXPECT_ANALYSIS_FINDING_PRESENT",
        "query" : {
          "type" : "ANALYZE_WORKBOOK_FINDINGS"
        },
        "code" : "FORMULA_ERROR_RESULT"
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ "Target rule: Matches the nested analysis query's target selectors." ]
  }
}
```

## EXPECT_ANALYSIS_FINDING_ABSENT

Run one analysis query and require no matching finding.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| query | REQUIRED | {"type":"TOP_LEVEL_TYPE_SET","typeSet":"inspectionQueryTypes"} |  |
| code | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| severity | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| messageContains | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "EXPECT_ANALYSIS_FINDING_ABSENT",
  "summary" : "Run one analysis query and require no matching finding.",
  "fields" : [ {
    "name" : "query",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "TOP_LEVEL_TYPE_SET",
      "typeSet" : "inspectionQueryTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "code",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "FORMULA_ERROR_RESULT", "FORMULA_EVALUATION_FAILURE", "FORMULA_EXTERNAL_REFERENCE", "FORMULA_MISSING_EXTERNAL_WORKBOOK", "FORMULA_USES_CACHED_EXTERNAL_VALUE", "FORMULA_UNREGISTERED_USER_DEFINED_FUNCTION", "FORMULA_VOLATILE_FUNCTION", "DATA_VALIDATION_EMPTY_EXPLICIT_LIST", "DATA_VALIDATION_MALFORMED_RULE", "DATA_VALIDATION_UNSUPPORTED_RULE", "DATA_VALIDATION_BROKEN_FORMULA", "DATA_VALIDATION_OVERLAPPING_RULES", "CONDITIONAL_FORMATTING_BROKEN_FORMULA", "CONDITIONAL_FORMATTING_UNSUPPORTED_RULE", "CONDITIONAL_FORMATTING_PRIORITY_COLLISION", "CONDITIONAL_FORMATTING_EMPTY_RANGE", "AUTOFILTER_INVALID_RANGE", "AUTOFILTER_MISSING_HEADER_ROW", "AUTOFILTER_TABLE_MISMATCH", "TABLE_OVERLAPPING_RANGE", "TABLE_BLANK_HEADER", "TABLE_DUPLICATE_HEADER", "TABLE_BROKEN_REFERENCE", "TABLE_STYLE_MISMATCH", "PIVOT_TABLE_MISSING_NAME", "PIVOT_TABLE_DUPLICATE_NAME", "PIVOT_TABLE_MISSING_CACHE_DEFINITION", "PIVOT_TABLE_MISSING_WORKBOOK_CACHE", "PIVOT_TABLE_MISSING_CACHE_RECORDS", "PIVOT_TABLE_BROKEN_SOURCE", "PIVOT_TABLE_UNSUPPORTED_DETAIL", "HYPERLINK_MALFORMED_TARGET", "HYPERLINK_MISSING_FILE_TARGET", "HYPERLINK_MISSING_DOCUMENT_SHEET", "HYPERLINK_UNRESOLVED_FILE_TARGET", "HYPERLINK_INVALID_DOCUMENT_TARGET", "NAMED_RANGE_BROKEN_REFERENCE", "NAMED_RANGE_UNPARSEABLE_FORMULA", "NAMED_RANGE_UNRESOLVED_TARGET", "NAMED_RANGE_SCOPE_SHADOWING" ]
  }, {
    "name" : "severity",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "INFO", "WARNING", "ERROR" ]
  }, {
    "name" : "messageContains",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ],
  "targetSelectorRule" : "Matches the nested analysis query's target selectors.",
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-expect-analysis-finding-absent",
      "assertion" : {
        "type" : "EXPECT_ANALYSIS_FINDING_ABSENT",
        "query" : {
          "type" : "ANALYZE_WORKBOOK_FINDINGS"
        },
        "code" : "FORMULA_ERROR_RESULT"
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ "Target rule: Matches the nested analysis query's target selectors." ]
  }
}
```

## ALL_OF

Require every nested assertion to pass against the same step target.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| assertions | REQUIRED | {"type":"LIST","elementShape":{"type":"TOP_LEVEL_TYPE_SET","typeSet":"assertionTypes"}} |  |

```json
{
  "id" : "ALL_OF",
  "summary" : "Require every nested assertion to pass against the same step target.",
  "fields" : [ {
    "name" : "assertions",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "TOP_LEVEL_TYPE_SET",
        "typeSet" : "assertionTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ],
  "targetSelectorRule" : "Matches the intersection of every nested assertion's target selectors.",
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-all-of",
      "assertion" : {
        "type" : "ALL_OF",
        "assertions" : [ {
          "type" : "EXPECT_ANALYSIS_MAX_SEVERITY",
          "query" : {
            "type" : "ANALYZE_WORKBOOK_FINDINGS"
          },
          "maximumSeverity" : "INFO"
        } ]
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ "Target rule: Matches the intersection of every nested assertion's target selectors." ]
  }
}
```

## ANY_OF

Require at least one nested assertion to pass against the same step target.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| assertions | REQUIRED | {"type":"LIST","elementShape":{"type":"TOP_LEVEL_TYPE_SET","typeSet":"assertionTypes"}} |  |

```json
{
  "id" : "ANY_OF",
  "summary" : "Require at least one nested assertion to pass against the same step target.",
  "fields" : [ {
    "name" : "assertions",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "TOP_LEVEL_TYPE_SET",
        "typeSet" : "assertionTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ],
  "targetSelectorRule" : "Matches the intersection of every nested assertion's target selectors.",
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-any-of",
      "assertion" : {
        "type" : "ANY_OF",
        "assertions" : [ {
          "type" : "EXPECT_ANALYSIS_MAX_SEVERITY",
          "query" : {
            "type" : "ANALYZE_WORKBOOK_FINDINGS"
          },
          "maximumSeverity" : "INFO"
        } ]
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ "Target rule: Matches the intersection of every nested assertion's target selectors." ]
  }
}
```

## NOT

Invert one nested assertion against the same step target.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| assertion | REQUIRED | {"type":"TOP_LEVEL_TYPE_SET","typeSet":"assertionTypes"} |  |

```json
{
  "id" : "NOT",
  "summary" : "Invert one nested assertion against the same step target.",
  "fields" : [ {
    "name" : "assertion",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "TOP_LEVEL_TYPE_SET",
      "typeSet" : "assertionTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ],
  "targetSelectorRule" : "Matches the nested assertion's target selectors.",
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "ASSERTION",
    "bodyField" : "assertion",
    "template" : {
      "stepId" : "step-not",
      "assertion" : {
        "type" : "NOT",
        "assertion" : {
          "type" : "EXPECT_ANALYSIS_MAX_SEVERITY",
          "query" : {
            "type" : "ANALYZE_WORKBOOK_FINDINGS"
          },
          "maximumSeverity" : "INFO"
        }
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ "Target rule: Matches the nested assertion's target selectors." ]
  }
}
```
