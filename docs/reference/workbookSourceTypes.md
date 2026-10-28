<!-- Generated from core metadata; do not edit. -->
# workbookSourceTypes

## NEW

Create a brand-new empty workbook. A new workbook starts with zero sheets; use ENSURE_SHEET to create the first sheet.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NEW",
  "summary" : "Create a brand-new empty workbook. A new workbook starts with zero sheets; use ENSURE_SHEET to create the first sheet.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## EXISTING

Open an existing .xlsx workbook from disk. source.security.passwordRef names a host-resolved secret for encrypted OOXML packages.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| path | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"PATH_SUFFIX","suffix":".xlsx"}]} |  |
| security | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"ooxmlOpenSecuritySpecType"} |  |

```json
{
  "id" : "EXISTING",
  "summary" : "Open an existing .xlsx workbook from disk. source.security.passwordRef names a host-resolved secret for encrypted OOXML packages.",
  "fields" : [ {
    "name" : "path",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "PATH_SUFFIX",
        "suffix" : ".xlsx"
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "security",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "ooxmlOpenSecuritySpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ],
  "noteRefs" : [ "requestOwnedPathRule" ]
}
```
