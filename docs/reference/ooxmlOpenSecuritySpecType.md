<!-- Generated from core metadata; do not edit. -->
# ooxmlOpenSecuritySpecType

## OoxmlOpenSecuritySpec

Optional OOXML package-open settings for encrypted existing workbook sources. password unlocks the encrypted OOXML package before GridGrind opens the inner .xlsx workbook.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| passwordRef | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"secretReferenceType"} |  |

```json
{
  "id" : "OoxmlOpenSecuritySpec",
  "summary" : "Optional OOXML package-open settings for encrypted existing workbook sources. password unlocks the encrypted OOXML package before GridGrind opens the inner .xlsx workbook.",
  "fields" : [ {
    "name" : "passwordRef",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "secretReferenceType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
