<!-- Generated from core metadata; do not edit. -->
# autofilterSortStateSpecType

## AutofilterSortStateSpec

Authored autofilter sort-state payload with one or more ordered sort conditions.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| caseSensitive | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| columnSort | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| sortMethod | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| conditions | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"autofilterSortConditionSpecTypes"}} |  |

```json
{
  "id" : "AutofilterSortStateSpec",
  "summary" : "Authored autofilter sort-state payload with one or more ordered sort conditions.",
  "fields" : [ {
    "name" : "range",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "caseSensitive",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "columnSort",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "sortMethod",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "PINYIN", "STROKE" ]
  }, {
    "name" : "conditions",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "autofilterSortConditionSpecTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
