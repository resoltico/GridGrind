<!-- Generated from core metadata; do not edit. -->
# ooxmlEncryptionReportTypes

## NONE

Package carries no OOXML encryption envelope.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Package carries no OOXML encryption envelope.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## ENCRYPTED

Package is encrypted with one fully specified OOXML encryption envelope.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| mode | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| cipherAlgorithm | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| hashAlgorithm | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| chainingMode | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| keyBits | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| blockSize | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| spinCount | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "ENCRYPTED",
  "summary" : "Package is encrypted with one fully specified OOXML encryption envelope.",
  "fields" : [ {
    "name" : "mode",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "AGILE", "STANDARD" ]
  }, {
    "name" : "cipherAlgorithm",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "RC4", "AES_128", "AES_192", "AES_256", "RC2", "DES", "TRIPLE_DES", "TRIPLE_DES_112", "RSA" ]
  }, {
    "name" : "hashAlgorithm",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "NONE", "SHA_1", "SHA_224", "SHA_256", "SHA_384", "SHA_512", "MD2", "MD4", "MD5", "RIPEMD_128", "RIPEMD_160", "RIPEMD_256", "WHIRLPOOL" ]
  }, {
    "name" : "chainingMode",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "ECB", "CBC", "CFB" ]
  }, {
    "name" : "keyBits",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "blockSize",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "spinCount",
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
