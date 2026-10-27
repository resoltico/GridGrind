<!-- Generated from core metadata; do not edit. -->
# workbookPersistenceTypes

## NONE

Keep the workbook in memory only. The response persistence.type echoes NONE.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "NONE",
  "summary" : "Keep the workbook in memory only. The response persistence.type echoes NONE.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## OVERWRITE

Overwrite the opened source workbook at source.path. No path field is accepted on OVERWRITE; the write target is the same path opened by the EXISTING source. persistence.security is required and declares both final OOXML encryption and signature dispositions. The response persistence.type echoes OVERWRITE, includes sourcePath (the original source path string) whenever an EXISTING source path was available, and otherwise omits sourcePath instead of inventing one. Its publication.status is NOT_ATTEMPTED, NOT_PUBLISHED, PUBLISHED or UNCERTAIN. PUBLISHED includes executionPath, SHA-256, byte size, staged-artifact verification and durability evidence; UNCERTAIN requires destination inspection before any retry.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| security | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"ooxmlPersistenceSecuritySpecType"} |  |

```json
{
  "id" : "OVERWRITE",
  "summary" : "Overwrite the opened source workbook at source.path. No path field is accepted on OVERWRITE; the write target is the same path opened by the EXISTING source. persistence.security is required and declares both final OOXML encryption and signature dispositions. The response persistence.type echoes OVERWRITE, includes sourcePath (the original source path string) whenever an EXISTING source path was available, and otherwise omits sourcePath instead of inventing one. Its publication.status is NOT_ATTEMPTED, NOT_PUBLISHED, PUBLISHED or UNCERTAIN. PUBLISHED includes executionPath, SHA-256, byte size, staged-artifact verification and durability evidence; UNCERTAIN requires destination inspection before any retry.",
  "fields" : [ {
    "name" : "security",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "ooxmlPersistenceSecuritySpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ],
  "noteRefs" : [ "requestOwnedPathRule" ]
}
```

## SAVE_AS

Save the workbook to one .xlsx path with an explicit ifExists collision policy. ifExists=REJECT requires the target path to be absent. ifExists=REPLACE enables create-or-replace. Existing-source writes require persistence.security to declare both final OOXML encryption and signature dispositions; NEW writes default to NONE on both axes when security is omitted. The response persistence.type echoes SAVE_AS, includes requestedPath (the literal path from the request), and carries publication.status as NOT_ATTEMPTED, NOT_PUBLISHED, PUBLISHED or UNCERTAIN. PUBLISHED includes executionPath, SHA-256, byte size, staged-artifact verification and durability evidence; UNCERTAIN requires destination inspection before any retry. Missing parent directories are created automatically.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| path | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"PATH_SUFFIX","suffix":".xlsx"}]} |  |
| ifExists | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| security | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"ooxmlPersistenceSecuritySpecType"} |  |

```json
{
  "id" : "SAVE_AS",
  "summary" : "Save the workbook to one .xlsx path with an explicit ifExists collision policy. ifExists=REJECT requires the target path to be absent. ifExists=REPLACE enables create-or-replace. Existing-source writes require persistence.security to declare both final OOXML encryption and signature dispositions; NEW writes default to NONE on both axes when security is omitted. The response persistence.type echoes SAVE_AS, includes requestedPath (the literal path from the request), and carries publication.status as NOT_ATTEMPTED, NOT_PUBLISHED, PUBLISHED or UNCERTAIN. PUBLISHED includes executionPath, SHA-256, byte size, staged-artifact verification and durability evidence; UNCERTAIN requires destination inspection before any retry. Missing parent directories are created automatically.",
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
    "name" : "ifExists",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "REJECT", "REPLACE" ]
  }, {
    "name" : "security",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "ooxmlPersistenceSecuritySpecType"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ],
  "noteRefs" : [ "requestOwnedPathRule" ]
}
```
