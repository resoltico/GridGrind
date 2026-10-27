<!-- Generated from core metadata; do not edit. -->
# customXmlExportReportType

## CustomXmlExportReport

One exported custom-XML mapping payload plus the factual mapping metadata used to produce it.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| mapping | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"customXmlMappingReportType"} |  |
| encoding | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| schemaValidated | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| xml | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "CustomXmlExportReport",
  "summary" : "One exported custom-XML mapping payload plus the factual mapping metadata used to produce it.",
  "fields" : [ {
    "name" : "mapping",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "customXmlMappingReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "encoding",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "schemaValidated",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "xml",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
