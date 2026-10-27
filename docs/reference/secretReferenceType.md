<!-- Generated from core metadata; do not edit. -->
# secretReferenceType

## SecretReference

An opaque identifier resolved by the host without carrying secret material.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| id | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"STRING_PATTERN","pattern":"[A-Za-z][A-Za-z0-9._-]*"}]} |  |

```json
{
  "id" : "SecretReference",
  "summary" : "An opaque identifier resolved by the host without carrying secret material.",
  "fields" : [ {
    "name" : "id",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "STRING_PATTERN",
        "pattern" : "[A-Za-z][A-Za-z0-9._-]*"
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
