<!-- Generated from core metadata; do not edit. -->
# tableRowSelectorTypes

## TABLE_ROW_ALL_ROWS

Select every logical data row in one selected table.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| table | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"tableSelectorTypes"} |  |

```json
{
  "id" : "TABLE_ROW_ALL_ROWS",
  "summary" : "Select every logical data row in one selected table.",
  "fields" : [ {
    "name" : "table",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "tableSelectorTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## TABLE_ROW_BY_INDEX

Select one zero-based data row by index in one selected table.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| table | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"tableSelectorTypes"} |  |
| rowIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "TABLE_ROW_BY_INDEX",
  "summary" : "Select one zero-based data row by index in one selected table.",
  "fields" : [ {
    "name" : "table",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "tableSelectorTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "rowIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## TABLE_ROW_BY_KEY_CELL

Select one logical data row by matching one key-column cell value.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| table | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"tableSelectorTypes"} |  |
| columnName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| expectedValue | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"cellValueSpecTypes"} |  |

```json
{
  "id" : "TABLE_ROW_BY_KEY_CELL",
  "summary" : "Select one logical data row by matching one key-column cell value.",
  "fields" : [ {
    "name" : "table",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "tableSelectorTypes"
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
  }, {
    "name" : "expectedValue",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellValueSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
