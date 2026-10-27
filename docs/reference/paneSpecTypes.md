<!-- Generated from core metadata; do not edit. -->
# paneSpecTypes

## NONE

Sheet has no active pane split or freeze.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Sheet has no active pane split or freeze.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## FROZEN

Freeze the sheet at the provided split and visible-origin coordinates.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| splitColumn | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| splitRow | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| leftmostColumn | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| topRow | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "FROZEN",
  "summary" : "Freeze the sheet at the provided split and visible-origin coordinates.",
  "fields" : [ {
    "name" : "splitColumn",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "splitRow",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "leftmostColumn",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "topRow",
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

## SPLIT

Apply split panes with explicit split offsets, visible origin, and active pane.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| xSplitPosition | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| ySplitPosition | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| leftmostColumn | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| topRow | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| activePane | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "SPLIT",
  "summary" : "Apply split panes with explicit split offsets, visible origin, and active pane.",
  "fields" : [ {
    "name" : "xSplitPosition",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "ySplitPosition",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "leftmostColumn",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "topRow",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "activePane",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "UPPER_LEFT", "UPPER_RIGHT", "LOWER_LEFT", "LOWER_RIGHT" ]
  } ],
  "targetSelectors" : [ ]
}
```
