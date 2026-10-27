<!-- Generated from core metadata; do not edit. -->
# customXmlMappingReportType

## CustomXmlMappingReport

One factual workbook custom-XML mapping report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| mapId | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| rootElement | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| schemaId | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| settings | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"customXmlMappingSettingsReportType"} |  |
| schema | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"customXmlMappingSchemaReportType"} |  |
| dataBinding | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"customXmlDataBindingReportType"} |  |
| linkedCells | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"customXmlLinkedCellReportType"}} |  |
| linkedTables | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"customXmlLinkedTableReportType"}} |  |

```json
{
  "id" : "CustomXmlMappingReport",
  "summary" : "One factual workbook custom-XML mapping report.",
  "fields" : [ {
    "name" : "mapId",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "rootElement",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "schemaId",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "settings",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "customXmlMappingSettingsReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "schema",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "customXmlMappingSchemaReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "dataBinding",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "customXmlDataBindingReportType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "linkedCells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "customXmlLinkedCellReportType"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "linkedTables",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "customXmlLinkedTableReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
