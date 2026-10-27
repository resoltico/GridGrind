<!-- Generated from core metadata; do not edit. -->
# cellValueSpecTypes

## BLANK

Write an empty cell.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |

```json
{
  "id" : "BLANK",
  "summary" : "Write an empty cell.",
  "fields" : [ ],
  "targetSelectors" : [ ]
}
```

## TEXT

Write one string cell value from an inline, file-backed, or standard-input text source. Blank resolved text is rejected; use BLANK for empty cells.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |

```json
{
  "id" : "TEXT",
  "summary" : "Write one string cell value from an inline, file-backed, or standard-input text source. Blank resolved text is rejected; use BLANK for empty cells.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## RICH_TEXT

Write a structured string cell value with an ordered rich-text run list. Run text concatenates to the stored plain string value and each run may override font attributes independently.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| runs | REQUIRED | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"richTextRunSpecType"}} |  |

```json
{
  "id" : "RICH_TEXT",
  "summary" : "Write a structured string cell value with an ordered rich-text run list. Run text concatenates to the stored plain string value and each run may override font attributes independently.",
  "fields" : [ {
    "name" : "runs",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "richTextRunSpecType"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## NUMBER

Write one numeric cell value.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"NUMBER"} |  |

```json
{
  "id" : "NUMBER",
  "summary" : "Write one numeric cell value.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "NUMBER"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## BOOLEAN

Write one boolean cell value.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"BOOLEAN"} |  |

```json
{
  "id" : "BOOLEAN",
  "summary" : "Write one boolean cell value.",
  "fields" : [ {
    "name" : "value",
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

## ERROR

Write one exact stored Excel error literal such as #REF!, #DIV/0!, or #N/A. Evaluation-only states such as #CIRCULAR_REF! are reported on readback but cannot be authored as stored cell values.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "ERROR",
  "summary" : "Write one exact stored Excel error literal such as #REF!, #DIV/0!, or #N/A. Evaluation-only states such as #CIRCULAR_REF! are reported on readback but cannot be authored as stored cell values.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A" ],
    "enumValueDocs" : [ {
      "value" : "#NULL!",
      "summary" : "Excel null-intersection error."
    }, {
      "value" : "#DIV/0!",
      "summary" : "Excel division-by-zero error."
    }, {
      "value" : "#VALUE!",
      "summary" : "Excel wrong-type or wrong-shape value error."
    }, {
      "value" : "#REF!",
      "summary" : "Excel invalid-reference error."
    }, {
      "value" : "#NAME?",
      "summary" : "Excel unknown-name error."
    }, {
      "value" : "#NUM!",
      "summary" : "Excel numeric-domain error."
    }, {
      "value" : "#N/A",
      "summary" : "Excel not-available error."
    } ]
  } ],
  "targetSelectors" : [ ]
}
```

## DATE

Write an ISO-8601 date value. Stored as an Excel serial number; readback keeps type=NUMBER, and the TEMPORAL facet can expose kind=DATE.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "DATE",
  "summary" : "Write an ISO-8601 date value. Stored as an Excel serial number; readback keeps type=NUMBER, and the TEMPORAL facet can expose kind=DATE.",
  "fields" : [ {
    "name" : "value",
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

## DATE_TIME

Write an ISO-8601 date-time value. Stored as an Excel serial number; readback keeps type=NUMBER, and the TEMPORAL facet can expose kind=DATE_TIME.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |

```json
{
  "id" : "DATE_TIME",
  "summary" : "Write an ISO-8601 date-time value. Stored as an Excel serial number; readback keeps type=NUMBER, and the TEMPORAL facet can expose kind=DATE_TIME.",
  "fields" : [ {
    "name" : "value",
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

## FORMULA

Write one Excel formula from an inline, file-backed, or standard-input text source. Formula text is the OOXML <f> body: it must not begin with = and is parsed by Apache POI.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |

```json
{
  "id" : "FORMULA",
  "summary" : "Write one Excel formula from an inline, file-backed, or standard-input text source. Formula text is the OOXML <f> body: it must not begin with = and is parsed by Apache POI.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## RAW_FORMULA

Write opaque OOXML <f> character data from one text source without POI formula parsing. It must be nonempty, must not begin with =, and rejects XML 1.0-invalid characters as INVALID_FORMULA_TEXT; XML-looking formula string literals are preserved.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| value | REQUIRED | {"type":"NESTED_TYPE_GROUP","group":"textSourceSpecTypes"} |  |

```json
{
  "id" : "RAW_FORMULA",
  "summary" : "Write opaque OOXML <f> character data from one text source without POI formula parsing. It must be nonempty, must not begin with =, and rejects XML 1.0-invalid characters as INVALID_FORMULA_TEXT; XML-looking formula string literals are preserved.",
  "fields" : [ {
    "name" : "value",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "textSourceSpecTypes"
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
