<!-- Generated from core metadata; do not edit. -->
# dataSourceTypes

## STRING_REFERENCE

Stored string reference and string-cache facts; formula absence/blankness is preserved.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| cache | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"stringCacheTypes"} |  |

```json
{
  "id" : "STRING_REFERENCE",
  "summary" : "Stored string reference and string-cache facts; formula absence/blankness is preserved.",
  "fields" : [ {
    "name" : "formula",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "cache",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "stringCacheTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## NUMERIC_REFERENCE

Stored numeric reference and numeric-cache facts; no values are substituted from live cells.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| cache | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"numericCacheTypes"} |  |

```json
{
  "id" : "NUMERIC_REFERENCE",
  "summary" : "Stored numeric reference and numeric-cache facts; no values are substituted from live cells.",
  "fields" : [ {
    "name" : "formula",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "cache",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "numericCacheTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## STRING_LITERAL

Stored literal string data, including declared count and physical point order.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| data | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"stringCacheTypes"} |  |

```json
{
  "id" : "STRING_LITERAL",
  "summary" : "Stored literal string data, including declared count and physical point order.",
  "fields" : [ {
    "name" : "data",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "stringCacheTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## NUMERIC_LITERAL

Stored literal numeric data, retaining raw value text, counts and formats.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| data | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"numericCacheTypes"} |  |

```json
{
  "id" : "NUMERIC_LITERAL",
  "summary" : "Stored literal numeric data, retaining raw value text, counts and formats.",
  "fields" : [ {
    "name" : "data",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "numericCacheTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## MISSING

No chart data-source element is stored.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "MISSING",
  "summary" : "No chart data-source element is stored.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## UNSUPPORTED

A stored chart data-source kind outside the typed observation contract; no substitute values are reported.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| nativeType | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| detail | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "UNSUPPORTED",
  "summary" : "A stored chart data-source kind outside the typed observation contract; no substitute values are reported.",
  "fields" : [ {
    "name" : "nativeType",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "detail",
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
