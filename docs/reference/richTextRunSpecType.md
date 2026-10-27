<!-- Generated from core metadata; do not edit. -->
# richTextRunSpecType

## RichTextRunSpec

One ordered rich-text run for a string cell. text must be non-empty; font is an optional override patch. The ordered run texts concatenate to the stored plain string value.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| text | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |
| font | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellFontPatchType"} |  |

```json
{
  "id" : "RichTextRunSpec",
  "summary" : "One ordered rich-text run for a string cell. text must be non-empty; font is an optional override patch. The ordered run texts concatenate to the stored plain string value.",
  "fields" : [ {
    "name" : "text",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "font",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellFontPatchType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
