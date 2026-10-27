<!-- Generated from core metadata; do not edit. -->
# hyperlinkTargetTypes

## URL

Attach an absolute URL target.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| target | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "URL",
  "summary" : "Attach an absolute URL target.",
  "fields" : [ {
    "name" : "target",
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

## EMAIL

Attach an email target without the mailto: prefix.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| email | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "EMAIL",
  "summary" : "Attach an email target without the mailto: prefix.",
  "fields" : [ {
    "name" : "email",
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

Attach a local or shared file path. Accepts plain paths or file: URIs and normalizes to a plain path string.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| path | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "FILE",
  "summary" : "Attach a local or shared file path. Accepts plain paths or file: URIs and normalizes to a plain path string.",
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
  "noteRefs" : [ "fileHyperlinkRelativePathRule" ]
}
```

## DOCUMENT

Attach an internal workbook target.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| target | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "DOCUMENT",
  "summary" : "Attach an internal workbook target.",
  "fields" : [ {
    "name" : "target",
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
