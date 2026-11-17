<!-- Generated from core metadata; do not edit. -->
# titleTypes

## NONE

No title is present.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "No title is present.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## TEXT

Static title text.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| text | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "TEXT",
  "summary" : "Static title text.",
  "fields" : [ {
    "name" : "text",
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

## FORMULA

Stored formula-backed title with its cache; never live-resolved text.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| cache | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"stringCacheTypes"} |  |

```json
{
  "id" : "FORMULA",
  "summary" : "Stored formula-backed title with its cache; never live-resolved text.",
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
