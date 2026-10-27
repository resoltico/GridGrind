<!-- Generated from core metadata; do not edit. -->
# chartTitleSpecTypes

## NONE

Remove any chart or series title.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Remove any chart or series title.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## TEXT

Use one explicit static title.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| source | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |

```json
{
  "id" : "TEXT",
  "summary" : "Use one explicit static title.",
  "fields" : [ {
    "name" : "source",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## FORMULA

Bind the chart or series title to one workbook formula that resolves to one cell.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "FORMULA",
  "summary" : "Bind the chart or series title to one workbook formula that resolves to one cell.",
  "fields" : [ {
    "name" : "formula",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
