<!-- Generated from core metadata; do not edit. -->
# printSetupSpecType

## PrintSetupSpec

Advanced page-setup payload nested under print-layout authoring. Supply the full authored setup block explicitly.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| margins | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"printMarginsSpecType"} |  |
| printGridlines | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| horizontallyCentered | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| verticallyCentered | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| paperSize | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| draft | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| blackAndWhite | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| copies | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| useFirstPageNumber | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| firstPageNumber | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |
| rowBreaks | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"NUMBER"}} |  |
| columnBreaks | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"NUMBER"}} |  |

```json
{
  "id" : "PrintSetupSpec",
  "summary" : "Advanced page-setup payload nested under print-layout authoring. Supply the full authored setup block explicitly.",
  "fields" : [ {
    "name" : "margins",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "printMarginsSpecType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "printGridlines",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "horizontallyCentered",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "verticallyCentered",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "paperSize",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "draft",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "blackAndWhite",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "copies",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "useFirstPageNumber",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "firstPageNumber",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  }, {
    "name" : "rowBreaks",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "NUMBER"
      }
    },
    "enumValues" : [ ]
  }, {
    "name" : "columnBreaks",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "NUMBER"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
