<!-- Generated from core metadata; do not edit. -->
# drawingCoordinateReportTypes

## EMU

Stored integer DrawingML coordinate in English metric units.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":-2.72730423296E13,"maximum":2.72730423169E13}]} |  |

```json
{
  "id" : "EMU",
  "summary" : "Stored integer DrawingML coordinate in English metric units.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "NUMBER_RANGE",
        "minimum" : -2.72730423296E13,
        "maximum" : 2.72730423169E13
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## MEASURED

Stored DrawingML measurement with its exact native units and lexical form.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| valueText | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"STRING_PATTERN","pattern":"-?[0-9]+(\\.[0-9]+)?(mm\|cm\|in\|pt\|pc\|pi)"}]} |  |

```json
{
  "id" : "MEASURED",
  "summary" : "Stored DrawingML measurement with its exact native units and lexical form.",
  "fields" : [ {
    "name" : "valueText",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "STRING_PATTERN",
        "pattern" : "-?[0-9]+(\\.[0-9]+)?(mm|cm|in|pt|pc|pi)"
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## PIXELS

Stored nonnegative VML client-anchor offset in pixels.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER","constraints":[{"type":"NUMBER_RANGE","minimum":0.0,"maximum":2.147483647E9}]} |  |

```json
{
  "id" : "PIXELS",
  "summary" : "Stored nonnegative VML client-anchor offset in pixels.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER",
      "constraints" : [ {
        "type" : "NUMBER_RANGE",
        "minimum" : 0.0,
        "maximum" : 2.147483647E9
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
