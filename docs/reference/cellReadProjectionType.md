<!-- Generated from core metadata; do not edit. -->
# cellReadProjectionType

## CellReadProjection

Shared readback projection that selects which factual cell facets are included on cell-returning inspection surfaces. Facets default to VALUE when omitted by the enclosing query. The facets field publishes enumValueDocs so each facet token explains the response fields it unlocks. Catalog field descriptors publish facet-gated response fields through projectedByFacets so agents can derive the required facet directly from the machine contract. Date-like numeric cells still read back as type=NUMBER until TEMPORAL is requested.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| facets | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "CellReadProjection",
  "summary" : "Shared readback projection that selects which factual cell facets are included on cell-returning inspection surfaces. Facets default to VALUE when omitted by the enclosing query. The facets field publishes enumValueDocs so each facet token explains the response fields it unlocks. Catalog field descriptors publish facet-gated response fields through projectedByFacets so agents can derive the required facet directly from the machine contract. Date-like numeric cells still read back as type=NUMBER until TEMPORAL is requested.",
  "fields" : [ {
    "name" : "facets",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "STRING"
      }
    },
    "enumValues" : [ "VALUE", "STYLE", "FORMAT", "HYPERLINK", "COMMENT", "FORMULA", "RICH_TEXT_RUNS", "TEMPORAL" ],
    "enumValueDocs" : [ {
      "value" : "VALUE",
      "summary" : "Project the factual effective scalar in value, including formula results and blank values."
    }, {
      "value" : "STYLE",
      "summary" : "Project the style report for each returned cell."
    }, {
      "value" : "FORMAT",
      "summary" : "Project displayValue using Excel's formatted display text."
    }, {
      "value" : "HYPERLINK",
      "summary" : "Project hyperlink metadata when the cell carries a hyperlink."
    }, {
      "value" : "COMMENT",
      "summary" : "Project comment metadata when the cell carries a comment."
    }, {
      "value" : "FORMULA",
      "summary" : "Project authored formula text for formula cells."
    }, {
      "value" : "RICH_TEXT_RUNS",
      "summary" : "Project rich-text runs for text cells and text-valued formula evaluations."
    }, {
      "value" : "TEMPORAL",
      "summary" : "Project derived date, time, or date-time semantics for date-like numeric values."
    } ]
  } ],
  "targetSelectors" : [ ]
}
```
