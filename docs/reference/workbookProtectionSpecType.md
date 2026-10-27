<!-- Generated from core metadata; do not edit. -->
# workbookProtectionSpecType

## WorkbookProtectionSpec

Workbook-protection payload covering workbook and revisions lock state plus optional passwords.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| structureLocked | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| windowsLocked | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| revisionsLocked | OPTIONAL | {"type":"SCALAR","scalarType":"BOOLEAN"} | false |
| workbookPasswordRef | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"secretReferenceType"} |  |
| revisionsPasswordRef | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"secretReferenceType"} |  |

```json
{
  "id" : "WorkbookProtectionSpec",
  "summary" : "Workbook-protection payload covering workbook and revisions lock state plus optional passwords.",
  "fields" : [ {
    "name" : "structureLocked",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "windowsLocked",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "revisionsLocked",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ],
    "defaultValue" : false
  }, {
    "name" : "workbookPasswordRef",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "secretReferenceType"
    },
    "enumValues" : [ ]
  }, {
    "name" : "revisionsPasswordRef",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "secretReferenceType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
