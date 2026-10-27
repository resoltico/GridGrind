<!-- Generated from core metadata; do not edit. -->
# customXmlLinkedTableReportType

## CustomXmlLinkedTableReport

One XML-mapped table linked to a custom-XML mapping.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| tableName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| tableDisplayName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| range | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| commonXPath | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "CustomXmlLinkedTableReport",
  "summary" : "One XML-mapped table linked to a custom-XML mapping.",
  "fields" : [ {
    "name" : "sheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "tableName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "tableDisplayName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "range",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "commonXPath",
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
