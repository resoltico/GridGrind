<!-- Generated from core metadata; do not edit. -->
# ooxmlSignatureSignerIdentityType

## OoxmlSignatureSignerIdentity

Signer identity material attached to one OOXML package signature report.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| subject | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| issuer | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| serialNumberHex | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "OoxmlSignatureSignerIdentity",
  "summary" : "Signer identity material attached to one OOXML package signature report.",
  "fields" : [ {
    "name" : "subject",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "issuer",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "serialNumberHex",
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
