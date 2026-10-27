<!-- Generated from core metadata; do not edit. -->
# colorSpecTypes

## RGB

Write one explicit RGB color reference plus optional tint metadata.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| rgb | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"STRING_PATTERN","pattern":"^#[0-9A-Fa-f]{6}$"}]} |  |
| tint | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "RGB",
  "summary" : "Write one explicit RGB color reference plus optional tint metadata.",
  "fields" : [ {
    "name" : "rgb",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "STRING_PATTERN",
        "pattern" : "^#[0-9A-Fa-f]{6}$"
      } ]
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

Write one workbook theme-slot color reference plus optional tint metadata.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| theme | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| tint | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "THEME",
  "summary" : "Write one workbook theme-slot color reference plus optional tint metadata.",
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

Write one indexed-palette color reference plus optional tint metadata.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| indexed | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| tint | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "INDEXED",
  "summary" : "Write one indexed-palette color reference plus optional tint metadata.",
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
