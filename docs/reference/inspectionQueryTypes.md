<!-- Generated from core metadata; do not edit. -->
# inspectionQueryTypes

## GET_WORKBOOK_SUMMARY

Return workbook-level summary facts.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=WorkbookSelector, typeIds=[WORKBOOK_CURRENT]]].

```json
{
  "id" : "GET_WORKBOOK_SUMMARY",
  "summary" : "Return workbook-level summary facts.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "WorkbookSelector",
    "typeIds" : [ "WORKBOOK_CURRENT" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-workbook-summary",
      "query" : {
        "type" : "GET_WORKBOOK_SUMMARY"
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ ]
  }
}
```

## GET_PACKAGE_SECURITY

Return OOXML package-encryption and package-signature facts.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=WorkbookSelector, typeIds=[WORKBOOK_CURRENT]]].

```json
{
  "id" : "GET_PACKAGE_SECURITY",
  "summary" : "Return OOXML package-encryption and package-signature facts.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "WorkbookSelector",
    "typeIds" : [ "WORKBOOK_CURRENT" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-package-security",
      "query" : {
        "type" : "GET_PACKAGE_SECURITY"
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ ]
  }
}
```

## GET_WORKBOOK_PROTECTION

Return workbook-level protection facts.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=WorkbookSelector, typeIds=[WORKBOOK_CURRENT]]].

```json
{
  "id" : "GET_WORKBOOK_PROTECTION",
  "summary" : "Return workbook-level protection facts.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "WorkbookSelector",
    "typeIds" : [ "WORKBOOK_CURRENT" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-workbook-protection",
      "query" : {
        "type" : "GET_WORKBOOK_PROTECTION"
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ ]
  }
}
```

## GET_CUSTOM_XML_MAPPINGS

Return workbook custom-XML mapping metadata.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=WorkbookSelector, typeIds=[WORKBOOK_CURRENT]]].

```json
{
  "id" : "GET_CUSTOM_XML_MAPPINGS",
  "summary" : "Return workbook custom-XML mapping metadata.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "WorkbookSelector",
    "typeIds" : [ "WORKBOOK_CURRENT" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-custom-xml-mappings",
      "query" : {
        "type" : "GET_CUSTOM_XML_MAPPINGS"
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ ]
  }
}
```

## EXPORT_CUSTOM_XML_MAPPING

Export one existing workbook custom-XML mapping as serialized XML.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| mapping | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"customXmlMappingLocatorType"} |  |
| validateSchema | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| encoding | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} | UTF-8 |

Targets: [TargetSelectorEntry[family=WorkbookSelector, typeIds=[WORKBOOK_CURRENT]]].

```json
{
  "id" : "EXPORT_CUSTOM_XML_MAPPING",
  "summary" : "Export one existing workbook custom-XML mapping as serialized XML.",
  "fields" : [ {
    "name" : "mapping",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "customXmlMappingLocatorType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "validateSchema",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "encoding",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "defaultValue" : "UTF-8"
  } ],
  "targetSelectors" : [ {
    "family" : "WorkbookSelector",
    "typeIds" : [ "WORKBOOK_CURRENT" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-export-custom-xml-mapping",
      "query" : {
        "type" : "EXPORT_CUSTOM_XML_MAPPING",
        "mapping" : {
          "name" : "Mapping1"
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

## GET_NAMED_RANGES

Return named ranges matched by the supplied selection.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=NamedRangeSelector, typeIds=[NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE]]].

```json
{
  "id" : "GET_NAMED_RANGES",
  "summary" : "Return named ranges matched by the supplied selection.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "NamedRangeSelector",
    "typeIds" : [ "NAMED_RANGE_ALL", "NAMED_RANGE_ANY_OF", "NAMED_RANGE_BY_NAMES", "NAMED_RANGE_BY_NAME", "NAMED_RANGE_WORKBOOK_SCOPE", "NAMED_RANGE_SHEET_SCOPE" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-named-ranges",
      "query" : {
        "type" : "GET_NAMED_RANGES"
      },
      "target" : {
        "type" : "NAMED_RANGE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_NAMED_RANGES accepts NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE." ]
  }
}
```

## GET_SHEET_SUMMARY

Return structural summary facts for one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "GET_SHEET_SUMMARY",
  "summary" : "Return structural summary facts for one sheet.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-sheet-summary",
      "query" : {
        "type" : "GET_SHEET_SUMMARY"
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

## GET_ARRAY_FORMULAS

Return factual array-formula group metadata for the selected sheets.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES]]].

```json
{
  "id" : "GET_ARRAY_FORMULAS",
  "summary" : "Return factual array-formula group metadata for the selected sheets.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_ALL", "SHEET_BY_NAME", "SHEET_BY_NAMES" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-array-formulas",
      "query" : {
        "type" : "GET_ARRAY_FORMULAS"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_ARRAY_FORMULAS accepts SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES." ]
  }
}
```

## GET_CELLS

Return exact cell snapshots for explicit addresses. Omit projection for the default compact VALUE readback.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| projection | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellReadProjectionType"} |  |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_BY_ADDRESS, CELL_BY_ADDRESSES]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "GET_CELLS",
  "summary" : "Return exact cell snapshots for explicit addresses. Omit projection for the default compact VALUE readback.",
  "fields" : [ {
    "name" : "projection",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellReadProjectionType"
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
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-cells",
      "query" : {
        "type" : "GET_CELLS"
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_CELLS accepts CELL_BY_ADDRESS, CELL_BY_ADDRESSES, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## GET_WINDOW

Return a rectangular window of cell snapshots. Omit projection for the default compact VALUE readback and omit includeBlanks for the sparse default.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| projection | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellReadProjectionType"} |  |
| includeBlanks | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_RECTANGULAR_WINDOW]]].

```json
{
  "id" : "GET_WINDOW",
  "summary" : "Return a rectangular window of cell snapshots. Omit projection for the default compact VALUE readback and omit includeBlanks for the sparse default.",
  "fields" : [ {
    "name" : "projection",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellReadProjectionType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "includeBlanks",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  } ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_RECTANGULAR_WINDOW" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-window",
      "query" : {
        "type" : "GET_WINDOW"
      },
      "target" : {
        "type" : "RANGE_RECTANGULAR_WINDOW",
        "sheetName" : "Sheet1",
        "topLeftAddress" : "A1",
        "rowCount" : 1,
        "columnCount" : 1
      }
    },
    "notes" : [ ]
  }
}
```

## GET_MERGED_REGIONS

Return the merged regions defined on one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "GET_MERGED_REGIONS",
  "summary" : "Return the merged regions defined on one sheet.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-merged-regions",
      "query" : {
        "type" : "GET_MERGED_REGIONS"
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

## GET_HYPERLINKS

Return hyperlink metadata for selected cells.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_ALL_USED_IN_SHEET, CELL_BY_ADDRESS, CELL_BY_ADDRESSES]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "GET_HYPERLINKS",
  "summary" : "Return hyperlink metadata for selected cells.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_ALL_USED_IN_SHEET", "CELL_BY_ADDRESS", "CELL_BY_ADDRESSES" ]
  }, {
    "family" : "TableCellSelector",
    "typeIds" : [ "TABLE_CELL_BY_COLUMN_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-hyperlinks",
      "query" : {
        "type" : "GET_HYPERLINKS"
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_HYPERLINKS accepts CELL_ALL_USED_IN_SHEET, CELL_BY_ADDRESS, CELL_BY_ADDRESSES, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## GET_COMMENTS

Return comment metadata for selected cells.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=CellSelector, typeIds=[CELL_ALL_USED_IN_SHEET, CELL_BY_ADDRESS, CELL_BY_ADDRESSES]], TargetSelectorEntry[family=TableCellSelector, typeIds=[TABLE_CELL_BY_COLUMN_NAME]]].

```json
{
  "id" : "GET_COMMENTS",
  "summary" : "Return comment metadata for selected cells.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "CellSelector",
    "typeIds" : [ "CELL_ALL_USED_IN_SHEET", "CELL_BY_ADDRESS", "CELL_BY_ADDRESSES" ]
  }, {
    "family" : "TableCellSelector",
    "typeIds" : [ "TABLE_CELL_BY_COLUMN_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-comments",
      "query" : {
        "type" : "GET_COMMENTS"
      },
      "target" : {
        "type" : "CELL_BY_ADDRESS",
        "sheetName" : "Sheet1",
        "address" : "A1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_COMMENTS accepts CELL_ALL_USED_IN_SHEET, CELL_BY_ADDRESS, CELL_BY_ADDRESSES, TABLE_CELL_BY_COLUMN_NAME." ]
  }
}
```

## GET_SHEET_LAYOUT

Return supported sheet-layout metadata for one sheet, including presentation, panes, tab color, defaults, ignored errors, and row or column outlineLevel state.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "GET_SHEET_LAYOUT",
  "summary" : "Return supported sheet-layout metadata for one sheet, including presentation, panes, tab color, defaults, ignored errors, and row or column outlineLevel state.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-sheet-layout",
      "query" : {
        "type" : "GET_SHEET_LAYOUT"
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

## GET_PRINT_LAYOUT

Return supported print-layout metadata for one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "GET_PRINT_LAYOUT",
  "summary" : "Return supported print-layout metadata for one sheet.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-print-layout",
      "query" : {
        "type" : "GET_PRINT_LAYOUT"
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

## GET_DATA_VALIDATIONS

Return factual data-validation structures for the selected sheet ranges.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_ALL_ON_SHEET, RANGE_BY_RANGE, RANGE_BY_RANGES, RANGE_RECTANGULAR_WINDOW]]].

```json
{
  "id" : "GET_DATA_VALIDATIONS",
  "summary" : "Return factual data-validation structures for the selected sheet ranges.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_ALL_ON_SHEET", "RANGE_BY_RANGE", "RANGE_BY_RANGES", "RANGE_RECTANGULAR_WINDOW" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-data-validations",
      "query" : {
        "type" : "GET_DATA_VALIDATIONS"
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_DATA_VALIDATIONS accepts RANGE_ALL_ON_SHEET, RANGE_BY_RANGE, RANGE_BY_RANGES, RANGE_RECTANGULAR_WINDOW." ]
  }
}
```

## GET_CONDITIONAL_FORMATTING

Return factual conditional-formatting blocks for the selected sheet ranges.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_ALL_ON_SHEET, RANGE_BY_RANGE, RANGE_BY_RANGES, RANGE_RECTANGULAR_WINDOW]]].

```json
{
  "id" : "GET_CONDITIONAL_FORMATTING",
  "summary" : "Return factual conditional-formatting blocks for the selected sheet ranges.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_ALL_ON_SHEET", "RANGE_BY_RANGE", "RANGE_BY_RANGES", "RANGE_RECTANGULAR_WINDOW" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-conditional-formatting",
      "query" : {
        "type" : "GET_CONDITIONAL_FORMATTING"
      },
      "target" : {
        "type" : "RANGE_BY_RANGE",
        "sheetName" : "Sheet1",
        "range" : "A1:B2"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_CONDITIONAL_FORMATTING accepts RANGE_ALL_ON_SHEET, RANGE_BY_RANGE, RANGE_BY_RANGES, RANGE_RECTANGULAR_WINDOW." ]
  }
}
```

## GET_AUTOFILTERS

Return sheet- and table-owned autofilter metadata for one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_BY_NAME]]].

```json
{
  "id" : "GET_AUTOFILTERS",
  "summary" : "Return sheet- and table-owned autofilter metadata for one sheet.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-autofilters",
      "query" : {
        "type" : "GET_AUTOFILTERS"
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

## GET_DRAWING_OBJECTS

Return factual drawing-object metadata for one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=DrawingObjectSelector, typeIds=[DRAWING_OBJECT_ALL_ON_SHEET]]].

```json
{
  "id" : "GET_DRAWING_OBJECTS",
  "summary" : "Return factual drawing-object metadata for one sheet.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "DrawingObjectSelector",
    "typeIds" : [ "DRAWING_OBJECT_ALL_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-drawing-objects",
      "query" : {
        "type" : "GET_DRAWING_OBJECTS"
      },
      "target" : {
        "type" : "DRAWING_OBJECT_ALL_ON_SHEET",
        "sheetName" : "Sheet1"
      }
    },
    "notes" : [ ]
  }
}
```

## GET_CHARTS

Return factual chart metadata for one sheet or one named chart.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=ChartSelector, typeIds=[CHART_ALL_ON_SHEET, CHART_BY_NAME]]].

```json
{
  "id" : "GET_CHARTS",
  "summary" : "Return factual chart metadata for one sheet or one named chart.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "ChartSelector",
    "typeIds" : [ "CHART_ALL_ON_SHEET", "CHART_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-charts",
      "query" : {
        "type" : "GET_CHARTS"
      },
      "target" : {
        "type" : "CHART_BY_NAME",
        "sheetName" : "Sheet1",
        "chartName" : "sample-chartName"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_CHARTS accepts CHART_ALL_ON_SHEET, CHART_BY_NAME." ]
  }
}
```

## GET_PIVOT_TABLES

Return factual pivot-table metadata selected by pivot-table name or ALL.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=PivotTableSelector, typeIds=[PIVOT_TABLE_ALL, PIVOT_TABLE_BY_NAME, PIVOT_TABLE_BY_NAMES, PIVOT_TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "GET_PIVOT_TABLES",
  "summary" : "Return factual pivot-table metadata selected by pivot-table name or ALL.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "PivotTableSelector",
    "typeIds" : [ "PIVOT_TABLE_ALL", "PIVOT_TABLE_BY_NAME", "PIVOT_TABLE_BY_NAMES", "PIVOT_TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-pivot-tables",
      "query" : {
        "type" : "GET_PIVOT_TABLES"
      },
      "target" : {
        "type" : "PIVOT_TABLE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_PIVOT_TABLES accepts PIVOT_TABLE_ALL, PIVOT_TABLE_BY_NAME, PIVOT_TABLE_BY_NAMES, PIVOT_TABLE_BY_NAME_ON_SHEET." ]
  }
}
```

## GET_DRAWING_OBJECT_PAYLOAD

Return the extracted binary payload for one named picture or embedded object.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=DrawingObjectSelector, typeIds=[DRAWING_OBJECT_BY_NAME]]].

```json
{
  "id" : "GET_DRAWING_OBJECT_PAYLOAD",
  "summary" : "Return the extracted binary payload for one named picture or embedded object.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "DrawingObjectSelector",
    "typeIds" : [ "DRAWING_OBJECT_BY_NAME" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-drawing-object-payload",
      "query" : {
        "type" : "GET_DRAWING_OBJECT_PAYLOAD"
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

## GET_TABLES

Return factual table metadata selected by workbook-global table name or ALL.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=TableSelector, typeIds=[TABLE_ALL, TABLE_BY_NAME, TABLE_BY_NAMES, TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "GET_TABLES",
  "summary" : "Return factual table metadata selected by workbook-global table name or ALL.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "TableSelector",
    "typeIds" : [ "TABLE_ALL", "TABLE_BY_NAME", "TABLE_BY_NAMES", "TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-tables",
      "query" : {
        "type" : "GET_TABLES"
      },
      "target" : {
        "type" : "TABLE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_TABLES accepts TABLE_ALL, TABLE_BY_NAME, TABLE_BY_NAMES, TABLE_BY_NAME_ON_SHEET." ]
  }
}
```

## GET_FORMULA_SURFACE

Return surface.totalFormulaCellCount plus per-sheet formula usage groups. surface.sheets[*] includes sheetName, formulaCellCount, distinctFormulaCount, and grouped formulas with occurrenceCount and addresses.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES]]].

```json
{
  "id" : "GET_FORMULA_SURFACE",
  "summary" : "Return surface.totalFormulaCellCount plus per-sheet formula usage groups. surface.sheets[*] includes sheetName, formulaCellCount, distinctFormulaCount, and grouped formulas with occurrenceCount and addresses.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_ALL", "SHEET_BY_NAME", "SHEET_BY_NAMES" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-formula-surface",
      "query" : {
        "type" : "GET_FORMULA_SURFACE"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_FORMULA_SURFACE accepts SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES." ]
  }
}
```

## GET_SHEET_SCHEMA

Infer a simple schema from a rectangular sheet window. Omit projection for the default compact VALUE readback.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| projection | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellReadProjectionType"} |  |

Targets: [TargetSelectorEntry[family=RangeSelector, typeIds=[RANGE_RECTANGULAR_WINDOW]]].

```json
{
  "id" : "GET_SHEET_SCHEMA",
  "summary" : "Infer a simple schema from a rectangular sheet window. Omit projection for the default compact VALUE readback.",
  "fields" : [ {
    "name" : "projection",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellReadProjectionType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ {
    "family" : "RangeSelector",
    "typeIds" : [ "RANGE_RECTANGULAR_WINDOW" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-sheet-schema",
      "query" : {
        "type" : "GET_SHEET_SCHEMA"
      },
      "target" : {
        "type" : "RANGE_RECTANGULAR_WINDOW",
        "sheetName" : "Sheet1",
        "topLeftAddress" : "A1",
        "rowCount" : 1,
        "columnCount" : 1
      }
    },
    "notes" : [ ]
  }
}
```

## GET_NAMED_RANGE_SURFACE

Return surface.workbookScopedCount, sheetScopedCount, rangeBackedCount, formulaBackedCount, and namedRanges. Each namedRanges entry reports name, scope, refersToFormula, and backing kind.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=NamedRangeSelector, typeIds=[NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE]]].

```json
{
  "id" : "GET_NAMED_RANGE_SURFACE",
  "summary" : "Return surface.workbookScopedCount, sheetScopedCount, rangeBackedCount, formulaBackedCount, and namedRanges. Each namedRanges entry reports name, scope, refersToFormula, and backing kind.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "NamedRangeSelector",
    "typeIds" : [ "NAMED_RANGE_ALL", "NAMED_RANGE_ANY_OF", "NAMED_RANGE_BY_NAMES", "NAMED_RANGE_BY_NAME", "NAMED_RANGE_WORKBOOK_SCOPE", "NAMED_RANGE_SHEET_SCOPE" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-get-named-range-surface",
      "query" : {
        "type" : "GET_NAMED_RANGE_SURFACE"
      },
      "target" : {
        "type" : "NAMED_RANGE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; GET_NAMED_RANGE_SURFACE accepts NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE." ]
  }
}
```

## ANALYZE_FORMULA_HEALTH

Return analysis.checkedFormulaCellCount, a severity summary, and findings for formula errors, volatile usage, or evaluation failures.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES]]].

```json
{
  "id" : "ANALYZE_FORMULA_HEALTH",
  "summary" : "Return analysis.checkedFormulaCellCount, a severity summary, and findings for formula errors, volatile usage, or evaluation failures.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_ALL", "SHEET_BY_NAME", "SHEET_BY_NAMES" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-analyze-formula-health",
      "query" : {
        "type" : "ANALYZE_FORMULA_HEALTH"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; ANALYZE_FORMULA_HEALTH accepts SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES." ]
  }
}
```

## ANALYZE_DATA_VALIDATION_HEALTH

Return data-validation-health findings for the selected sheets.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES]]].

```json
{
  "id" : "ANALYZE_DATA_VALIDATION_HEALTH",
  "summary" : "Return data-validation-health findings for the selected sheets.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_ALL", "SHEET_BY_NAME", "SHEET_BY_NAMES" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-analyze-data-validation-health",
      "query" : {
        "type" : "ANALYZE_DATA_VALIDATION_HEALTH"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; ANALYZE_DATA_VALIDATION_HEALTH accepts SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES." ]
  }
}
```

## ANALYZE_CONDITIONAL_FORMATTING_HEALTH

Return conditional-formatting-health findings for the selected sheets.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES]]].

```json
{
  "id" : "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
  "summary" : "Return conditional-formatting-health findings for the selected sheets.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_ALL", "SHEET_BY_NAME", "SHEET_BY_NAMES" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-analyze-conditional-formatting-health",
      "query" : {
        "type" : "ANALYZE_CONDITIONAL_FORMATTING_HEALTH"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; ANALYZE_CONDITIONAL_FORMATTING_HEALTH accepts SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES." ]
  }
}
```

## ANALYZE_AUTOFILTER_HEALTH

Return autofilter-health findings for the selected sheets.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES]]].

```json
{
  "id" : "ANALYZE_AUTOFILTER_HEALTH",
  "summary" : "Return autofilter-health findings for the selected sheets.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_ALL", "SHEET_BY_NAME", "SHEET_BY_NAMES" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-analyze-autofilter-health",
      "query" : {
        "type" : "ANALYZE_AUTOFILTER_HEALTH"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; ANALYZE_AUTOFILTER_HEALTH accepts SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES." ]
  }
}
```

## ANALYZE_TABLE_HEALTH

Return table-health findings for the selected tables.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=TableSelector, typeIds=[TABLE_ALL, TABLE_BY_NAME, TABLE_BY_NAMES, TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "ANALYZE_TABLE_HEALTH",
  "summary" : "Return table-health findings for the selected tables.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "TableSelector",
    "typeIds" : [ "TABLE_ALL", "TABLE_BY_NAME", "TABLE_BY_NAMES", "TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-analyze-table-health",
      "query" : {
        "type" : "ANALYZE_TABLE_HEALTH"
      },
      "target" : {
        "type" : "TABLE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; ANALYZE_TABLE_HEALTH accepts TABLE_ALL, TABLE_BY_NAME, TABLE_BY_NAMES, TABLE_BY_NAME_ON_SHEET." ]
  }
}
```

## ANALYZE_PIVOT_TABLE_HEALTH

Return pivot-table-health findings for the selected pivot tables.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=PivotTableSelector, typeIds=[PIVOT_TABLE_ALL, PIVOT_TABLE_BY_NAME, PIVOT_TABLE_BY_NAMES, PIVOT_TABLE_BY_NAME_ON_SHEET]]].

```json
{
  "id" : "ANALYZE_PIVOT_TABLE_HEALTH",
  "summary" : "Return pivot-table-health findings for the selected pivot tables.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "PivotTableSelector",
    "typeIds" : [ "PIVOT_TABLE_ALL", "PIVOT_TABLE_BY_NAME", "PIVOT_TABLE_BY_NAMES", "PIVOT_TABLE_BY_NAME_ON_SHEET" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-analyze-pivot-table-health",
      "query" : {
        "type" : "ANALYZE_PIVOT_TABLE_HEALTH"
      },
      "target" : {
        "type" : "PIVOT_TABLE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; ANALYZE_PIVOT_TABLE_HEALTH accepts PIVOT_TABLE_ALL, PIVOT_TABLE_BY_NAME, PIVOT_TABLE_BY_NAMES, PIVOT_TABLE_BY_NAME_ON_SHEET." ]
  }
}
```

## ANALYZE_HYPERLINK_HEALTH

Return hyperlink-health findings for the selected sheets.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=SheetSelector, typeIds=[SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES]]].

```json
{
  "id" : "ANALYZE_HYPERLINK_HEALTH",
  "summary" : "Return hyperlink-health findings for the selected sheets.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "SheetSelector",
    "typeIds" : [ "SHEET_ALL", "SHEET_BY_NAME", "SHEET_BY_NAMES" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-analyze-hyperlink-health",
      "query" : {
        "type" : "ANALYZE_HYPERLINK_HEALTH"
      },
      "target" : {
        "type" : "SHEET_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; ANALYZE_HYPERLINK_HEALTH accepts SHEET_ALL, SHEET_BY_NAME, SHEET_BY_NAMES." ]
  }
}
```

## ANALYZE_NAMED_RANGE_HEALTH

Return analysis.checkedNamedRangeCount, a severity summary, and named-range findings such as broken references, unresolved targets, or scope shadowing.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=NamedRangeSelector, typeIds=[NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE]]].

```json
{
  "id" : "ANALYZE_NAMED_RANGE_HEALTH",
  "summary" : "Return analysis.checkedNamedRangeCount, a severity summary, and named-range findings such as broken references, unresolved targets, or scope shadowing.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "NamedRangeSelector",
    "typeIds" : [ "NAMED_RANGE_ALL", "NAMED_RANGE_ANY_OF", "NAMED_RANGE_BY_NAMES", "NAMED_RANGE_BY_NAME", "NAMED_RANGE_WORKBOOK_SCOPE", "NAMED_RANGE_SHEET_SCOPE" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-analyze-named-range-health",
      "query" : {
        "type" : "ANALYZE_NAMED_RANGE_HEALTH"
      },
      "target" : {
        "type" : "NAMED_RANGE_BY_NAME",
        "name" : "Sheet1"
      }
    },
    "notes" : [ "Replace the placeholder target when needed; ANALYZE_NAMED_RANGE_HEALTH accepts NAMED_RANGE_ALL, NAMED_RANGE_ANY_OF, NAMED_RANGE_BY_NAMES, NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE, NAMED_RANGE_SHEET_SCOPE." ]
  }
}
```

## ANALYZE_WORKBOOK_FINDINGS

Return analysis.summary plus one flat analysis.findings list after running all analysis families (formula health, data-validation health, conditional-formatting health, autofilter health, table health, pivot-table health, hyperlink health, named-range health) across the entire workbook and aggregate findings in a single response. This is the primary workbook-health check and pairs naturally with persistence.type=NONE when no save is required.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

Targets: [TargetSelectorEntry[family=WorkbookSelector, typeIds=[WORKBOOK_CURRENT]]].

```json
{
  "id" : "ANALYZE_WORKBOOK_FINDINGS",
  "summary" : "Return analysis.summary plus one flat analysis.findings list after running all analysis families (formula health, data-validation health, conditional-formatting health, autofilter health, table health, pivot-table health, hyperlink health, named-range health) across the entire workbook and aggregate findings in a single response. This is the primary workbook-health check and pairs naturally with persistence.type=NONE when no save is required.",
  "fields" : [ ],
  "targetSelectors" : [ {
    "family" : "WorkbookSelector",
    "typeIds" : [ "WORKBOOK_CURRENT" ]
  } ],
  "effects" : [ "READ_WORKBOOK" ],
  "effectFootprint" : "TARGET_BOUNDED",
  "stepTemplate" : {
    "stepKind" : "INSPECTION",
    "bodyField" : "query",
    "template" : {
      "stepId" : "step-analyze-workbook-findings",
      "query" : {
        "type" : "ANALYZE_WORKBOOK_FINDINGS"
      },
      "target" : {
        "type" : "WORKBOOK_CURRENT"
      }
    },
    "notes" : [ ]
  }
}
```
