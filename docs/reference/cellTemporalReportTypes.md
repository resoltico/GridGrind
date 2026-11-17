<!-- Generated from core metadata; do not edit. -->
# cellTemporalReportTypes

## NOT_TEMPORAL

The observed format does not denote a temporal value.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NOT_TEMPORAL",
  "summary" : "The observed format does not denote a temporal value.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## INTERPRETED

The numeric fact has a supported civil DATE, TIME or DATE_TIME interpretation.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| kind | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| isoValue | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "INTERPRETED",
  "summary" : "The numeric fact has a supported civil DATE, TIME or DATE_TIME interpretation.",
  "fields" : [ {
    "name" : "kind",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "DATE", "TIME", "DATE_TIME", "ELAPSED" ]
  }, {
    "name" : "isoValue",
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

## UNREPRESENTABLE

The format is temporal, but its numeric fact has no supported interpretation; request VALUE for the original number.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| kind | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| reason | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "UNREPRESENTABLE",
  "summary" : "The format is temporal, but its numeric fact has no supported interpretation; request VALUE for the original number.",
  "fields" : [ {
    "name" : "kind",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "DATE", "TIME", "DATE_TIME", "ELAPSED" ]
  }, {
    "name" : "reason",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "NONFINITE", "NEGATIVE", "OUTSIDE_CIVIL_RANGE", "FICTITIOUS_1900_LEAP_DAY", "UNSUPPORTED_ELAPSED_FORMAT" ]
  } ],
  "targetSelectors" : [ ]
}
```
