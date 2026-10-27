<!-- Generated from core metadata; do not edit. -->
# commentReportType

## CommentReport

Exact factual comment report including optional rich-text runs and anchor bounds.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| text | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| author | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| visible | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| runs | OPTIONAL | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"richTextRunReportType"}} |  |
| anchor | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"commentAnchorReportType"} |  |

```json
{
  "id" : "CommentReport",
  "summary" : "Exact factual comment report including optional rich-text runs and anchor bounds.",
  "fields" : [ {
    "name" : "text",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "author",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "visible",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "runs",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "richTextRunReportType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "anchor",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "commentAnchorReportType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
