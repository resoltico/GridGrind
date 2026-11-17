<!-- Generated from core metadata; do not edit. -->
# axisCrossingTypes

## MISSING

No axis crossing declaration is stored.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "MISSING",
  "summary" : "No axis crossing declaration is stored.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## MODE

Stored axis crossing mode.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| mode | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "MODE",
  "summary" : "Stored axis crossing mode.",
  "fields" : [ {
    "name" : "mode",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "AUTO_ZERO", "MAX", "MIN" ]
  } ],
  "targetSelectors" : [ ]
}
```

## VALUE

Stored numeric axis crossing with its exact native attribute text.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| valueText | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "VALUE",
  "summary" : "Stored numeric axis crossing with its exact native attribute text.",
  "fields" : [ {
    "name" : "valueText",
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
