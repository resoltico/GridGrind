<!-- Generated from core metadata; do not edit. -->
# chartDataSourceSpecTypes

## REFERENCE

Workbook-backed chart source formula or defined name.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "REFERENCE",
  "summary" : "Workbook-backed chart source formula or defined name.",
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
| values | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"NUMBER"}} |  |

```json
{
  "id" : "NUMERIC_LITERAL",
  "summary" : "Literal numeric chart source stored directly in the chart part.",
  "fields" : [ {
    "name" : "values",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "NUMBER"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
