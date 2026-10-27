<!-- Generated from core metadata; do not edit. -->
# printScalingSpecTypes

## AUTOMATIC

Sheet uses Excel's default scaling instead of fit-to-page counts.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "AUTOMATIC",
  "summary" : "Sheet uses Excel's default scaling instead of fit-to-page counts.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## FIT

Sheet fits printed content into the provided page counts. A value of 0 on one axis keeps that axis unconstrained.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| widthPages | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| heightPages | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "FIT",
  "summary" : "Sheet fits printed content into the provided page counts. A value of 0 on one axis keeps that axis unconstrained.",
  "fields" : [ {
    "name" : "widthPages",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "heightPages",
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
