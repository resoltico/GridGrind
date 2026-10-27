<!-- Generated from core metadata; do not edit. -->
# workbookStepTypes

## MUTATION

Execute one mutation action against the selected workbook target.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| stepId | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"},{"type":"STRING_PATTERN","pattern":"[A-Za-z0-9._-]+"}]} |  |
| target | REQUIRED | {"type":"NESTED_TYPE_GROUP_UNION","groups":["workbookSelectorTypes","sheetSelectorTypes","cellSelectorTypes","rangeSelectorTypes","rowBandSelectorTypes","columnBandSelectorTypes","tableSelectorTypes","tableRowSelectorTypes","tableCellSelectorTypes","namedRangeSelectorTypes","drawingObjectSelectorTypes","chartSelectorTypes","pivotTableSelectorTypes"]} |  |
| action | REQUIRED | {"type":"TOP_LEVEL_TYPE_SET","typeSet":"mutationActionTypes"} |  |

```json
{
  "id" : "MUTATION",
  "summary" : "Execute one mutation action against the selected workbook target.",
  "fields" : [ {
    "name" : "stepId",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      }, {
        "type" : "STRING_PATTERN",
        "pattern" : "[A-Za-z0-9._-]+"
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "target",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP_UNION",
      "groups" : [ "workbookSelectorTypes", "sheetSelectorTypes", "cellSelectorTypes", "rangeSelectorTypes", "rowBandSelectorTypes", "columnBandSelectorTypes", "tableSelectorTypes", "tableRowSelectorTypes", "tableCellSelectorTypes", "namedRangeSelectorTypes", "drawingObjectSelectorTypes", "chartSelectorTypes", "pivotTableSelectorTypes" ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "action",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "TOP_LEVEL_TYPE_SET",
      "typeSet" : "mutationActionTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## ASSERTION

Verify one authored expectation against the selected workbook target.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| stepId | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"},{"type":"STRING_PATTERN","pattern":"[A-Za-z0-9._-]+"}]} |  |
| target | REQUIRED | {"type":"NESTED_TYPE_GROUP_UNION","groups":["workbookSelectorTypes","sheetSelectorTypes","cellSelectorTypes","rangeSelectorTypes","rowBandSelectorTypes","columnBandSelectorTypes","tableSelectorTypes","tableRowSelectorTypes","tableCellSelectorTypes","namedRangeSelectorTypes","drawingObjectSelectorTypes","chartSelectorTypes","pivotTableSelectorTypes"]} |  |
| assertion | REQUIRED | {"type":"TOP_LEVEL_TYPE_SET","typeSet":"assertionTypes"} |  |

```json
{
  "id" : "ASSERTION",
  "summary" : "Verify one authored expectation against the selected workbook target.",
  "fields" : [ {
    "name" : "stepId",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      }, {
        "type" : "STRING_PATTERN",
        "pattern" : "[A-Za-z0-9._-]+"
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "target",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP_UNION",
      "groups" : [ "workbookSelectorTypes", "sheetSelectorTypes", "cellSelectorTypes", "rangeSelectorTypes", "rowBandSelectorTypes", "columnBandSelectorTypes", "tableSelectorTypes", "tableRowSelectorTypes", "tableCellSelectorTypes", "namedRangeSelectorTypes", "drawingObjectSelectorTypes", "chartSelectorTypes", "pivotTableSelectorTypes" ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "assertion",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "TOP_LEVEL_TYPE_SET",
      "typeSet" : "assertionTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## INSPECTION

Run one factual or analytical inspection query against the selected workbook target.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| stepId | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"},{"type":"STRING_PATTERN","pattern":"[A-Za-z0-9._-]+"}]} |  |
| target | REQUIRED | {"type":"NESTED_TYPE_GROUP_UNION","groups":["workbookSelectorTypes","sheetSelectorTypes","cellSelectorTypes","rangeSelectorTypes","rowBandSelectorTypes","columnBandSelectorTypes","tableSelectorTypes","tableRowSelectorTypes","tableCellSelectorTypes","namedRangeSelectorTypes","drawingObjectSelectorTypes","chartSelectorTypes","pivotTableSelectorTypes"]} |  |
| query | REQUIRED | {"type":"TOP_LEVEL_TYPE_SET","typeSet":"inspectionQueryTypes"} |  |

```json
{
  "id" : "INSPECTION",
  "summary" : "Run one factual or analytical inspection query against the selected workbook target.",
  "fields" : [ {
    "name" : "stepId",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      }, {
        "type" : "STRING_PATTERN",
        "pattern" : "[A-Za-z0-9._-]+"
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "target",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP_UNION",
      "groups" : [ "workbookSelectorTypes", "sheetSelectorTypes", "cellSelectorTypes", "rangeSelectorTypes", "rowBandSelectorTypes", "columnBandSelectorTypes", "tableSelectorTypes", "tableRowSelectorTypes", "tableCellSelectorTypes", "namedRangeSelectorTypes", "drawingObjectSelectorTypes", "chartSelectorTypes", "pivotTableSelectorTypes" ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "query",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "TOP_LEVEL_TYPE_SET",
      "typeSet" : "inspectionQueryTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
