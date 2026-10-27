<!-- Generated from core metadata; do not edit. -->
# binarySourceSpecTypes

## INLINE_BASE64

Embed base64-encoded binary content directly in the request JSON.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| base64Data | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "INLINE_BASE64",
  "summary" : "Embed base64-encoded binary content directly in the request JSON.",
  "fields" : [ {
    "name" : "base64Data",
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

## FILE

Load binary content from one file path.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| path | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "FILE",
  "summary" : "Load binary content from one file path.",
  "fields" : [ {
    "name" : "path",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ],
  "noteRefs" : [ "requestOwnedPathRule" ]
}
```

## STANDARD_INPUT

Load binary content from the execution transport's bound standard input bytes.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "STANDARD_INPUT",
  "summary" : "Load binary content from the execution transport's bound standard input bytes.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```
