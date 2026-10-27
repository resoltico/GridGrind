<!-- Generated from core metadata; do not edit. -->
# customXmlDataBindingReportType

## CustomXmlDataBindingReport

Optional custom-XML data-binding metadata attached to one workbook mapping.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| dataBindingName | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| fileBinding | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| connectionId | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| fileBindingName | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| loadMode | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "CustomXmlDataBindingReport",
  "summary" : "Optional custom-XML data-binding metadata attached to one workbook mapping.",
  "fields" : [ {
    "name" : "dataBindingName",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "fileBinding",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "connectionId",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "fileBindingName",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "loadMode",
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
