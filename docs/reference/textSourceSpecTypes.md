<!-- Generated from core metadata; do not edit. -->
# textSourceSpecTypes

## INLINE

Embed UTF-8 text directly in the request JSON.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| text | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "INLINE",
  "summary" : "Embed UTF-8 text directly in the request JSON.",
  "fields" : [ {
    "name" : "text",
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

## UTF8_FILE

Load UTF-8 text from one file path.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| path | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "UTF8_FILE",
  "summary" : "Load UTF-8 text from one file path.",
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

Load UTF-8 text from the execution transport's bound standard input bytes.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "STANDARD_INPUT",
  "summary" : "Load UTF-8 text from the execution transport's bound standard input bytes.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```
