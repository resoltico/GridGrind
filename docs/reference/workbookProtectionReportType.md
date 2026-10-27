<!-- Generated from core metadata; do not edit. -->
# workbookProtectionReportType

## WorkbookProtectionReport

Exact workbook-protection report covering structure, windows, revisions, and password-hash presence flags.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| structureLocked | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| windowsLocked | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| revisionsLocked | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| workbookPasswordHashPresent | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |
| revisionsPasswordHashPresent | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "WorkbookProtectionReport",
  "summary" : "Exact workbook-protection report covering structure, windows, revisions, and password-hash presence flags.",
  "fields" : [ {
    "name" : "structureLocked",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "windowsLocked",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "revisionsLocked",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "workbookPasswordHashPresent",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  }, {
    "name" : "revisionsPasswordHashPresent",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "BOOLEAN"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
