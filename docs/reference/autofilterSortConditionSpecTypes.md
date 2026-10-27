<!-- Generated from core metadata; do not edit. -->
# autofilterSortConditionSpecTypes

## VALUE

Sort by the ordinary cell value with no auxiliary discriminator payload.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| descending | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "VALUE",
  "summary" : "Sort by the ordinary cell value with no auxiliary discriminator payload.",
  "fields" : [ {
    "name" : "range",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "descending",
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

## CELL_COLOR

Sort by the cell fill color referenced by one explicit color.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| descending | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| color | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |

```json
{
  "id" : "CELL_COLOR",
  "summary" : "Sort by the cell fill color referenced by one explicit color.",
  "fields" : [ {
    "name" : "range",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "descending",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "color",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## FONT_COLOR

Sort by the rendered font color referenced by one explicit color.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| descending | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| color | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |

```json
{
  "id" : "FONT_COLOR",
  "summary" : "Sort by the rendered font color referenced by one explicit color.",
  "fields" : [ {
    "name" : "range",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "descending",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "color",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "colorSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## ICON

Sort by the icon id inside the active icon-set definition.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| descending | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| iconId | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "ICON",
  "summary" : "Sort by the icon id inside the active icon-set definition.",
  "fields" : [ {
    "name" : "range",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "descending",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "iconId",
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
