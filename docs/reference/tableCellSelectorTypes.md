<!-- Generated from core metadata; do not edit. -->
# tableCellSelectorTypes

## TABLE_CELL_BY_COLUMN_NAME

Select one logical cell within one selected table row by column name.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| row | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"tableRowSelectorTypes"} |  |
| columnName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "TABLE_CELL_BY_COLUMN_NAME",
  "summary" : "Select one logical cell within one selected table row by column name.",
  "fields" : [ {
    "name" : "row",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "tableRowSelectorTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "columnName",
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
