<!-- Generated from core metadata; do not edit. -->
# sheetProtectionReportTypes

## UNPROTECTED

Expect the sheet to have no protection.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "UNPROTECTED",
  "summary" : "Expect the sheet to have no protection.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## PROTECTED

Expect the sheet to be protected with explicit lock settings.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| settings | REQUIRED | {"type":"PLAIN_TYPE_GROUP","group":"sheetProtectionSettingsType"} |  |

```json
{
  "id" : "PROTECTED",
  "summary" : "Expect the sheet to be protected with explicit lock settings.",
  "fields" : [ {
    "name" : "settings",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "sheetProtectionSettingsType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
