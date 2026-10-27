<!-- Generated from core metadata; do not edit. -->
# ooxmlPersistenceEncryptionSpecTypes

## NONE

Deliberately persist plaintext.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Deliberately persist plaintext.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## ENCRYPT

Persist with one explicitly supplied OOXML encryption envelope.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| encryption | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"ooxmlEncryptionSpecType"} |  |

```json
{
  "id" : "ENCRYPT",
  "summary" : "Persist with one explicitly supplied OOXML encryption envelope.",
  "fields" : [ {
    "name" : "encryption",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "ooxmlEncryptionSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## PRESERVE_SOURCE

Reapply the verified source encryption envelope when it is compatible with the AGILE write contract.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "PRESERVE_SOURCE",
  "summary" : "Reapply the verified source encryption envelope when it is compatible with the AGILE write contract.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```
