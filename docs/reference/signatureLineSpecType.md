<!-- Generated from core metadata; do not edit. -->
# signatureLineSpecType

## SignatureLineSpec

Named signature-line authoring payload for SET_SIGNATURE_LINE. allowComments is explicit and plainSignature is optional, but caption or suggested signer metadata must still be present.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| name | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| anchor | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"drawingAnchorSpecTypes"} |  |
| allowComments | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| signingInstructions | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| suggestedSigner | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| suggestedSigner2 | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| suggestedSignerEmail | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| caption | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| invalidStamp | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| plainSignature | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"pictureDataSpecType"} |  |

```json
{
  "id" : "SignatureLineSpec",
  "summary" : "Named signature-line authoring payload for SET_SIGNATURE_LINE. allowComments is explicit and plainSignature is optional, but caption or suggested signer metadata must still be present.",
  "fields" : [ {
    "name" : "name",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "anchor",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "drawingAnchorSpecTypes"
    },
    "enumValues" : [ ]
  }, {
    "name" : "allowComments",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "signingInstructions",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "suggestedSigner",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "suggestedSigner2",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "suggestedSignerEmail",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "caption",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "invalidStamp",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "plainSignature",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "pictureDataSpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
