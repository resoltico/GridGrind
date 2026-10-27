<!-- Generated from core metadata; do not edit. -->
# cellReportTypes

## BLANK

Factual blank cell report. Only address and type are always present; VALUE projects the blank scalar, and FORMAT, STYLE, HYPERLINK, and COMMENT project their metadata.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| address | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| displayValue | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| style | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellStyleReportType"} |  |
| hyperlink | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"hyperlinkTargetTypes"} |  |
| comment | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"commentReportType"} |  |
| value | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"cellValueTypes"} |  |

```json
{
  "id" : "BLANK",
  "summary" : "Factual blank cell report. Only address and type are always present; VALUE projects the blank scalar, and FORMAT, STYLE, HYPERLINK, and COMMENT project their metadata.",
  "fields" : [ {
    "name" : "address",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "displayValue",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "FORMAT" ]
  }, {
    "name" : "style",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellStyleReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "STYLE" ]
  }, {
    "name" : "hyperlink",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "hyperlinkTargetTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "HYPERLINK" ]
  }, {
    "name" : "comment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "commentReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "COMMENT" ]
  }, {
    "name" : "value",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellValueTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "VALUE" ]
  } ],
  "targetSelectors" : [ ]
}
```

## TEXT

Factual text cell report. VALUE projects value, RICH_TEXT_RUNS projects runs, and the remaining fields are facet-gated.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| address | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| displayValue | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| style | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellStyleReportType"} |  |
| hyperlink | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"hyperlinkTargetTypes"} |  |
| comment | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"commentReportType"} |  |
| value | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"cellValueTypes"} |  |
| runs | OPTIONAL | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"richTextRunReportType"}} |  |

```json
{
  "id" : "TEXT",
  "summary" : "Factual text cell report. VALUE projects value, RICH_TEXT_RUNS projects runs, and the remaining fields are facet-gated.",
  "fields" : [ {
    "name" : "address",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "displayValue",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "FORMAT" ]
  }, {
    "name" : "style",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellStyleReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "STYLE" ]
  }, {
    "name" : "hyperlink",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "hyperlinkTargetTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "HYPERLINK" ]
  }, {
    "name" : "comment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "commentReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "COMMENT" ]
  }, {
    "name" : "value",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellValueTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "VALUE" ]
  }, {
    "name" : "runs",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "richTextRunReportType"
      }
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "RICH_TEXT_RUNS" ]
  } ],
  "targetSelectors" : [ ]
}
```

## NUMBER

Factual numeric cell report. VALUE projects value; TEMPORAL adds date, time, or date-time semantics when a numeric Excel value is format-derived rather than a plain number.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| address | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| displayValue | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| style | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellStyleReportType"} |  |
| hyperlink | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"hyperlinkTargetTypes"} |  |
| comment | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"commentReportType"} |  |
| value | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"cellValueTypes"} |  |
| temporal | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellTemporalReportType"} |  |

```json
{
  "id" : "NUMBER",
  "summary" : "Factual numeric cell report. VALUE projects value; TEMPORAL adds date, time, or date-time semantics when a numeric Excel value is format-derived rather than a plain number.",
  "fields" : [ {
    "name" : "address",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "displayValue",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "FORMAT" ]
  }, {
    "name" : "style",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellStyleReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "STYLE" ]
  }, {
    "name" : "hyperlink",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "hyperlinkTargetTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "HYPERLINK" ]
  }, {
    "name" : "comment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "commentReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "COMMENT" ]
  }, {
    "name" : "value",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellValueTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "VALUE" ]
  }, {
    "name" : "temporal",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellTemporalReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "TEMPORAL" ]
  } ],
  "targetSelectors" : [ ]
}
```

## BOOLEAN

Factual boolean cell report. VALUE projects value and the remaining fields are facet-gated.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| address | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| displayValue | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| style | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellStyleReportType"} |  |
| hyperlink | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"hyperlinkTargetTypes"} |  |
| comment | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"commentReportType"} |  |
| value | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"cellValueTypes"} |  |

```json
{
  "id" : "BOOLEAN",
  "summary" : "Factual boolean cell report. VALUE projects value and the remaining fields are facet-gated.",
  "fields" : [ {
    "name" : "address",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "displayValue",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "FORMAT" ]
  }, {
    "name" : "style",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellStyleReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "STYLE" ]
  }, {
    "name" : "hyperlink",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "hyperlinkTargetTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "HYPERLINK" ]
  }, {
    "name" : "comment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "commentReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "COMMENT" ]
  }, {
    "name" : "value",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellValueTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "VALUE" ]
  } ],
  "targetSelectors" : [ ]
}
```

## ERROR

Factual error cell report. VALUE projects value and the remaining fields are facet-gated.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| address | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| displayValue | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| style | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellStyleReportType"} |  |
| hyperlink | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"hyperlinkTargetTypes"} |  |
| comment | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"commentReportType"} |  |
| value | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"cellValueTypes"} |  |

```json
{
  "id" : "ERROR",
  "summary" : "Factual error cell report. VALUE projects value and the remaining fields are facet-gated.",
  "fields" : [ {
    "name" : "address",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "displayValue",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "FORMAT" ]
  }, {
    "name" : "style",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellStyleReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "STYLE" ]
  }, {
    "name" : "hyperlink",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "hyperlinkTargetTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "HYPERLINK" ]
  }, {
    "name" : "comment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "commentReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "COMMENT" ]
  }, {
    "name" : "value",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellValueTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "VALUE" ]
  } ],
  "targetSelectors" : [ ]
}
```

## FORMULA

Factual formula cell report with separate projected scalar value. FORMULA projects formula text, VALUE projects value, and the remaining fields are facet-gated.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| address | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| displayValue | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| style | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellStyleReportType"} |  |
| hyperlink | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"hyperlinkTargetTypes"} |  |
| comment | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"commentReportType"} |  |
| formula | OPTIONAL | {"type":"SCALAR","scalarType":"STRING"} |  |
| value | OPTIONAL | {"type":"NESTED_TYPE_GROUP","group":"cellValueTypes"} |  |
| runs | OPTIONAL | {"type":"LIST","elementShape":{"type":"PLAIN_TYPE_GROUP","group":"richTextRunReportType"}} |  |
| temporal | OPTIONAL | {"type":"PLAIN_TYPE_GROUP","group":"cellTemporalReportType"} |  |

```json
{
  "id" : "FORMULA",
  "summary" : "Factual formula cell report with separate projected scalar value. FORMULA projects formula text, VALUE projects value, and the remaining fields are facet-gated.",
  "fields" : [ {
    "name" : "address",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ]
  }, {
    "name" : "displayValue",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "FORMAT" ]
  }, {
    "name" : "style",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellStyleReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "STYLE" ]
  }, {
    "name" : "hyperlink",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "hyperlinkTargetTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "HYPERLINK" ]
  }, {
    "name" : "comment",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "commentReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "COMMENT" ]
  }, {
    "name" : "formula",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "FORMULA" ]
  }, {
    "name" : "value",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "NESTED_TYPE_GROUP",
      "group" : "cellValueTypes"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "VALUE" ]
  }, {
    "name" : "runs",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "PLAIN_TYPE_GROUP",
        "group" : "richTextRunReportType"
      }
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "RICH_TEXT_RUNS" ]
  }, {
    "name" : "temporal",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "PLAIN_TYPE_GROUP",
      "group" : "cellTemporalReportType"
    },
    "enumValues" : [ ],
    "projectedByFacets" : [ "TEMPORAL" ]
  } ],
  "targetSelectors" : [ ]
}
```
