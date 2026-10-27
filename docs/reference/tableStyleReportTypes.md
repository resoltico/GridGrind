<!-- Generated from core metadata; do not edit. -->
# tableStyleReportTypes

## NONE

Expect the table to carry no persisted style.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Expect the table to carry no persisted style.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## NAMED

Expect the table to carry one named style plus stripe/emphasis flags.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| showFirstColumn | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| showLastColumn | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| showRowStripes | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| showColumnStripes | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "NAMED",
  "summary" : "Expect the table to carry one named style plus stripe/emphasis flags.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "showFirstColumn",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "showLastColumn",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "showRowStripes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "showColumnStripes",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
