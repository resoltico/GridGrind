<!-- Generated from core metadata; do not edit. -->
# ooxmlPersistenceSecuritySpecType

## OoxmlPersistenceSecuritySpec

Total OOXML encryption and signature policy applied during persistence. It is required for writing EXISTING sources; NEW writes default to encryption NONE and signature NONE when security is omitted.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| encryption | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"ooxmlPersistenceEncryptionSpecTypes"} |  |
| signature | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"ooxmlPersistenceSignatureSpecTypes"} |  |

```json
{
  "id" : "OoxmlPersistenceSecuritySpec",
  "summary" : "Total OOXML encryption and signature policy applied during persistence. It is required for writing EXISTING sources; NEW writes default to encryption NONE and signature NONE when security is omitted.",
  "fields" : [ {
    "name" : "encryption",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "ooxmlPersistenceEncryptionSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "signature",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "ooxmlPersistenceSignatureSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
