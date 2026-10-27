<!-- Generated from core metadata; do not edit. -->
# cellColorReportTypes

## RGB

Read one explicit RGB workbook color plus optional tint metadata.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| rgb | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| tint | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "RGB",
  "summary" : "Read one explicit RGB workbook color plus optional tint metadata.",
  "fields" : [ {
    "name" : "rgb",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "tint",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## THEME

Read one workbook theme-slot color plus optional tint metadata.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| theme | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| tint | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "THEME",
  "summary" : "Read one workbook theme-slot color plus optional tint metadata.",
  "fields" : [ {
    "name" : "theme",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "tint",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## INDEXED

Read one indexed-palette workbook color plus optional tint metadata.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| indexed | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| tint | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "INDEXED",
  "summary" : "Read one indexed-palette workbook color plus optional tint metadata.",
  "fields" : [ {
    "name" : "indexed",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "tint",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
