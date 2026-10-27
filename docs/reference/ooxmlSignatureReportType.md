<!-- Generated from core metadata; do not edit. -->
# ooxmlSignatureReportType

## OoxmlSignatureReport

Factual OOXML package-signature report for one signature part. state reflects the currently loaded workbook package, including INVALIDATED_BY_MUTATION for source signatures after in-memory edits.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| packagePartName | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| signer | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"ooxmlSignatureSignerIdentityType"} |  |
| state | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "OoxmlSignatureReport",
  "summary" : "Factual OOXML package-signature report for one signature part. state reflects the currently loaded workbook package, including INVALIDATED_BY_MUTATION for source signatures after in-memory edits.",
  "fields" : [ {
    "name" : "packagePartName",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "signer",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "ooxmlSignatureSignerIdentityType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "state",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "VALID", "INVALID", "INVALIDATED_BY_MUTATION" ]
  } ],
  "targetSelectors" : [ ]
}
```
