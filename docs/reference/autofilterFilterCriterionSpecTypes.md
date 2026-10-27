<!-- Generated from core metadata; do not edit. -->
# autofilterFilterCriterionSpecTypes

## VALUES

Retain rows whose cell values match one or more explicit values.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| values | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |
| includeBlank | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "VALUES",
  "summary" : "Retain rows whose cell values match one or more explicit values.",
  "fields" : [ {
    "name" : "values",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "STRING"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "includeBlank",
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

## CUSTOM

Retain rows that satisfy one or two comparator-based custom conditions.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| and | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| conditions | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"autofilterCustomConditionInputType"}} |  |

```json
{
  "id" : "CUSTOM",
  "summary" : "Retain rows that satisfy one or two comparator-based custom conditions.",
  "fields" : [ {
    "name" : "and",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "conditions",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "autofilterCustomConditionInputType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## DYNAMIC

Retain rows using one dynamic-date or moving-window autofilter rule.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| dynamicType | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| value | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| maxValue | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "DYNAMIC",
  "summary" : "Retain rows using one dynamic-date or moving-window autofilter rule.",
  "fields" : [ {
    "name" : "dynamicType",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "value",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "maxValue",
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

## TOP10

Retain top or bottom N or percent values.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| top | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| percent | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "TOP10",
  "summary" : "Retain top or bottom N or percent values.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "top",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "percent",
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

## COLOR

Retain rows matching one cell color or font color criterion.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cellColor | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| color | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"colorSpecTypes"} |  |

```json
{
  "id" : "COLOR",
  "summary" : "Retain rows matching one cell color or font color criterion.",
  "fields" : [ {
    "name" : "cellColor",
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

Retain rows matching one icon-set member.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| iconSet | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| iconId | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "ICON",
  "summary" : "Retain rows matching one icon-set member.",
  "fields" : [ {
    "name" : "iconSet",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
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
