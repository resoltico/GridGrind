<!-- Generated from core metadata; do not edit. -->
# ooxmlSignatureSpecType

## OoxmlSignatureSpec

OOXML package-signing settings for workbook persistence. Omit keyPasswordRef to reuse the secret resolved by keystorePasswordRef. digestAlgorithm defaults to SHA256 when omitted. alias may be omitted only when the keystore contains exactly one signable private-key entry.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| pkcs12Path | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| keystorePasswordRef | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"secretReferenceType"} |  |
| keyPasswordRef | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"secretReferenceType"} |  |
| alias | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| digestAlgorithm | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} | SHA256 |
| description | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "OoxmlSignatureSpec",
  "summary" : "OOXML package-signing settings for workbook persistence. Omit keyPasswordRef to reuse the secret resolved by keystorePasswordRef. digestAlgorithm defaults to SHA256 when omitted. alias may be omitted only when the keystore contains exactly one signable private-key entry.",
  "fields" : [ {
    "name" : "pkcs12Path",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "keystorePasswordRef",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "secretReferenceType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "keyPasswordRef",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "secretReferenceType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "alias",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "digestAlgorithm",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "SHA256", "SHA384", "SHA512" ],
    "defaultValue" : "SHA256"
  }, {
    "name" : "description",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ],
  "noteRefs" : [ "requestOwnedPathRule" ]
}
```
