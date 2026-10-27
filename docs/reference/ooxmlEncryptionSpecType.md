<!-- Generated from core metadata; do not edit. -->
# ooxmlEncryptionSpecType

## OoxmlEncryptionSpec

OOXML package-encryption settings for workbook persistence. GridGrind writes AGILE packages only. cipher defaults to AES_256 and hash defaults to SHA_512 when omitted. Supported ciphers: AES_256 and AES_192. Supported hashes: SHA_512, SHA_384, and SHA_256 (LIM-038).

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| passwordRef | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"secretReferenceType"} |  |
| cipher | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} | AES_256 |
| hash | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} | SHA_512 |

```json
{
  "id" : "OoxmlEncryptionSpec",
  "summary" : "OOXML package-encryption settings for workbook persistence. GridGrind writes AGILE packages only. cipher defaults to AES_256 and hash defaults to SHA_512 when omitted. Supported ciphers: AES_256 and AES_192. Supported hashes: SHA_512, SHA_384, and SHA_256 (LIM-038).",
  "fields" : [ {
    "name" : "passwordRef",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "secretReferenceType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "cipher",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "AES_256", "AES_192" ],
    "defaultValue" : "AES_256"
  }, {
    "name" : "hash",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "SHA_512", "SHA_384", "SHA_256" ],
    "defaultValue" : "SHA_512"
  } ],
  "targetSelectors" : [ ]
}
```
