<!-- Generated from core metadata; do not edit. -->
# conditionalFormattingThresholdSpecTypes

## MIN

Use the minimum value observed in the applied range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "MIN",
  "summary" : "Use the minimum value observed in the applied range.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## MAX

Use the maximum value observed in the applied range.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "MAX",
  "summary" : "Use the maximum value observed in the applied range.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## NUMBER

Use one finite numeric threshold.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "NUMBER",
  "summary" : "Use one finite numeric threshold.",
  "fields" : [ {
    "name" : "value",
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

## PERCENT

Use one finite percentage threshold from 0 through 100.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "PERCENT",
  "summary" : "Use one finite percentage threshold from 0 through 100.",
  "fields" : [ {
    "name" : "value",
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

## PERCENTILE

Use one finite percentile threshold from 0 through 100.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "PERCENTILE",
  "summary" : "Use one finite percentile threshold from 0 through 100.",
  "fields" : [ {
    "name" : "value",
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

## FORMULA

Use one nonblank, DDE-safe formula threshold.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |

```json
{
  "id" : "FORMULA",
  "summary" : "Use one nonblank, DDE-safe formula threshold.",
  "fields" : [ {
    "name" : "formula",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
