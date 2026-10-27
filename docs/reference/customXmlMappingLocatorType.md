<!-- Generated from core metadata; do not edit. -->
# customXmlMappingLocatorType

## CustomXmlMappingLocator

One locator for an existing workbook custom-XML mapping. Supply mapId, name, or both; the locator must resolve to exactly one existing mapping.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| mapId | OPTIONAL | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| name | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "CustomXmlMappingLocator",
  "summary" : "One locator for an existing workbook custom-XML mapping. Supply mapId, name, or both; the locator must resolve to exactly one existing mapping.",
  "fields" : [ {
    "name" : "mapId",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "name",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
