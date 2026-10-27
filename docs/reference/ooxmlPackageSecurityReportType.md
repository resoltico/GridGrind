<!-- Generated from core metadata; do not edit. -->
# ooxmlPackageSecurityReportType

## OoxmlPackageSecurityReport

Factual OOXML package-security report covering encryption and package signatures.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| encryption | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"ooxmlEncryptionReportTypes"} |  |
| signatures | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"ooxmlSignatureReportType"}} |  |

```json
{
  "id" : "OoxmlPackageSecurityReport",
  "summary" : "Factual OOXML package-security report covering encryption and package signatures.",
  "fields" : [ {
    "name" : "encryption",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "ooxmlEncryptionReportTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "signatures",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "ooxmlSignatureReportType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
