<!-- Generated from core metadata; do not edit. -->
# namedRangeSelectorTypes

## NAMED_RANGE_ALL

Select every user-facing named range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NAMED_RANGE_ALL",
  "summary" : "Select every user-facing named range.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## NAMED_RANGE_ANY_OF

Select the union of one or more explicit named-range references.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| selectors | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"refTypes"}} |  |

```json
{
  "id" : "NAMED_RANGE_ANY_OF",
  "summary" : "Select the union of one or more explicit named-range references.",
  "fields" : [ {
    "name" : "selectors",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "refTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## NAMED_RANGE_BY_NAMES

Match named ranges across all scopes by exact name set.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| names | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "NAMED_RANGE_BY_NAMES",
  "summary" : "Match named ranges across all scopes by exact name set.",
  "fields" : [ {
    "name" : "names",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "STRING"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## NAMED_RANGE_BY_NAME

Match a named range reference across all scopes by exact name.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "NAMED_RANGE_BY_NAME",
  "summary" : "Match a named range reference across all scopes by exact name.",
  "fields" : [ {
    "name" : "name",
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

## NAMED_RANGE_WORKBOOK_SCOPE

Match one workbook-scoped named range reference by exact name.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "NAMED_RANGE_WORKBOOK_SCOPE",
  "summary" : "Match one workbook-scoped named range reference by exact name.",
  "fields" : [ {
    "name" : "name",
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

## NAMED_RANGE_SHEET_SCOPE

Match one sheet-scoped named range reference on one sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "NAMED_RANGE_SHEET_SCOPE",
  "summary" : "Match one sheet-scoped named range reference on one sheet.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "sheetName",
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
