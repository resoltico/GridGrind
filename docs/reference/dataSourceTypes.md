<!-- Generated from core metadata; do not edit. -->
# dataSourceTypes

## STRING_REFERENCE

Formula-backed string chart source plus cached values.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| cachedValues | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "STRING_REFERENCE",
  "summary" : "Formula-backed string chart source plus cached values.",
  "fields" : [ {
    "name" : "formula",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "cachedValues",
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

## NUMERIC_REFERENCE

Formula-backed numeric chart source plus cached values.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| formatCode | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| cachedValues | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "NUMERIC_REFERENCE",
  "summary" : "Formula-backed numeric chart source plus cached values.",
  "fields" : [ {
    "name" : "formula",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "formatCode",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "cachedValues",
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

## STRING_LITERAL

Literal string chart source stored directly in the chart part.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| values | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "STRING_LITERAL",
  "summary" : "Literal string chart source stored directly in the chart part.",
  "fields" : [ {
    "name" : "values",
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

## NUMERIC_LITERAL

Literal numeric chart source stored directly in the chart part.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formatCode | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| values | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "NUMERIC_LITERAL",
  "summary" : "Literal numeric chart source stored directly in the chart part.",
  "fields" : [ {
    "name" : "formatCode",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "values",
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
