<!-- Generated from core metadata; do not edit. -->
# conditionalFormattingSpecType

## ConditionalFormattingSpec

One authored conditional-formatting definition with an ordered rule list. Target ranges are owned by the selector; rules must not be empty.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| rules | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"conditionalFormattingRuleSpecTypes"}} |  |

```json
{
  "id" : "ConditionalFormattingSpec",
  "summary" : "One authored conditional-formatting definition with an ordered rule list. Target ranges are owned by the selector; rules must not be empty.",
  "fields" : [ {
    "name" : "rules",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "conditionalFormattingRuleSpecTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
