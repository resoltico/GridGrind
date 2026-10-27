<!-- Generated from core metadata; do not edit. -->
# customXmlImportSpecType

## CustomXmlImportSpec

One custom-XML import payload targeting an existing workbook mapping plus the XML content to import.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| locator | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"customXmlMappingLocatorType"} |  |
| xml | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |

```json
{
  "id" : "CustomXmlImportSpec",
  "summary" : "One custom-XML import payload targeting an existing workbook mapping plus the XML content to import.",
  "fields" : [ {
    "name" : "locator",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "customXmlMappingLocatorType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "xml",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
