<!-- Generated from core metadata; do not edit. -->
# richTextRunReportType

## RichTextRunReport

One factual rich-text run reported from a text cell or comment.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| text | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| font | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"cellFontReportType"} |  |

```json
{
  "id" : "RichTextRunReport",
  "summary" : "One factual rich-text run reported from a text cell or comment.",
  "fields" : [ {
    "name" : "text",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "font",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellFontReportType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
