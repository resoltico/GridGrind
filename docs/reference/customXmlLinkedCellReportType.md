<!-- Generated from core metadata; do not edit. -->
# customXmlLinkedCellReportType

## CustomXmlLinkedCellReport

One single-cell binding linked to a custom-XML mapping.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| sheetName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| address | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| xpath | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| xmlDataType | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "CustomXmlLinkedCellReport",
  "summary" : "One single-cell binding linked to a custom-XML mapping.",
  "fields" : [ {
    "name" : "sheetName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "address",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "xpath",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "xmlDataType",
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
