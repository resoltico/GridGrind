<!-- Generated from core metadata; do not edit. -->
# ooxmlPersistenceSignatureSpecTypes

## NONE

Deliberately persist without an OOXML package signature.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Deliberately persist without an OOXML package signature.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## SIGN

Sign with explicitly supplied PKCS#12 material.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| signature | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"ooxmlSignatureSpecType"} |  |

```json
{
  "id" : "SIGN",
  "summary" : "Sign with explicitly supplied PKCS#12 material.",
  "fields" : [ {
    "name" : "signature",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "ooxmlSignatureSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
