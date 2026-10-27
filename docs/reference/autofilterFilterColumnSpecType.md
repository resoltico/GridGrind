<!-- Generated from core metadata; do not edit. -->
# autofilterFilterColumnSpecType

## AutofilterFilterColumnSpec

One authored autofilter filter-column payload with an explicit column criterion.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| columnId | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| showButton | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | true |
| criterion | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"autofilterFilterCriterionSpecTypes"} |  |

```json
{
  "id" : "AutofilterFilterColumnSpec",
  "summary" : "One authored autofilter filter-column payload with an explicit column criterion.",
  "fields" : [ {
    "name" : "columnId",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "showButton",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : true
  }, {
    "name" : "criterion",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "autofilterFilterCriterionSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
