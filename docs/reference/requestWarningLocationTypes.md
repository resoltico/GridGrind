<!-- Generated from core metadata; do not edit. -->
# requestWarningLocationTypes

## STEP

Warning located at one authored workbook step.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| stepIndex | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| stepId | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| stepType | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "STEP",
  "summary" : "Warning located at one authored workbook step.",
  "fields" : [ {
    "name" : "stepIndex",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "stepId",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "stepType",
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

## REQUEST_PATH

Warning located at one request-owned filesystem path.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| path | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| pathRole | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "REQUEST_PATH",
  "summary" : "Warning located at one request-owned filesystem path.",
  "fields" : [ {
    "name" : "path",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "pathRole",
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

## REQUEST_BYTE_OFFSET

Warning located at one authored request byte offset.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| byteOffset | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "REQUEST_BYTE_OFFSET",
  "summary" : "Warning located at one authored request byte offset.",
  "fields" : [ {
    "name" : "byteOffset",
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

## FORMULA_CELL

Warning located at one formula cell assessed by calculation capability analysis.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| address | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "FORMULA_CELL",
  "summary" : "Warning located at one formula cell assessed by calculation capability analysis.",
  "fields" : [ {
    "name" : "sheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "address",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
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
