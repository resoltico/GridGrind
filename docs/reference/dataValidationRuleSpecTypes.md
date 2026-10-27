<!-- Generated from core metadata; do not edit. -->
# dataValidationRuleSpecTypes

## EXPLICIT_LIST

Allow only one of the supplied explicit values. An empty values array preserves Excel's explicit-empty-list state.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| values | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "EXPLICIT_LIST",
  "summary" : "Allow only one of the supplied explicit values. An empty values array preserves Excel's explicit-empty-list state.",
  "fields" : [ {
    "name" : "values",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "STRING"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## FORMULA_LIST

Allow values from a formula-driven list expression.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |

```json
{
  "id" : "FORMULA_LIST",
  "summary" : "Allow values from a formula-driven list expression.",
  "fields" : [ {
    "name" : "formula",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## WHOLE_NUMBER

Apply a whole-number comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| operator | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| formula1 | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |
| formula2 | OPTIONAL | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |

```json
{
  "id" : "WHOLE_NUMBER",
  "summary" : "Apply a whole-number comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.",
  "fields" : [ {
    "name" : "operator",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "BETWEEN", "NOT_BETWEEN", "EQUAL", "NOT_EQUAL", "GREATER_THAN", "LESS_THAN", "GREATER_OR_EQUAL", "LESS_OR_EQUAL" ]
  }, {
    "name" : "formula1",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "formula2",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## DECIMAL_NUMBER

Apply a decimal-number comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| operator | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| formula1 | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |
| formula2 | OPTIONAL | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |

```json
{
  "id" : "DECIMAL_NUMBER",
  "summary" : "Apply a decimal-number comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.",
  "fields" : [ {
    "name" : "operator",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "BETWEEN", "NOT_BETWEEN", "EQUAL", "NOT_EQUAL", "GREATER_THAN", "LESS_THAN", "GREATER_OR_EQUAL", "LESS_OR_EQUAL" ]
  }, {
    "name" : "formula1",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "formula2",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## DATE

Apply a date comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| operator | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| formula1 | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |
| formula2 | OPTIONAL | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |

```json
{
  "id" : "DATE",
  "summary" : "Apply a date comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.",
  "fields" : [ {
    "name" : "operator",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "BETWEEN", "NOT_BETWEEN", "EQUAL", "NOT_EQUAL", "GREATER_THAN", "LESS_THAN", "GREATER_OR_EQUAL", "LESS_OR_EQUAL" ]
  }, {
    "name" : "formula1",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "formula2",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## TIME

Apply a time comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| operator | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| formula1 | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |
| formula2 | OPTIONAL | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |

```json
{
  "id" : "TIME",
  "summary" : "Apply a time comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.",
  "fields" : [ {
    "name" : "operator",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "BETWEEN", "NOT_BETWEEN", "EQUAL", "NOT_EQUAL", "GREATER_THAN", "LESS_THAN", "GREATER_OR_EQUAL", "LESS_OR_EQUAL" ]
  }, {
    "name" : "formula1",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "formula2",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## TEXT_LENGTH

Apply a text-length comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| operator | REQUIRED | {"type":"SCALAR","scalarType":"STRING"} |  |
| formula1 | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |
| formula2 | OPTIONAL | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |

```json
{
  "id" : "TEXT_LENGTH",
  "summary" : "Apply a text-length comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.",
  "fields" : [ {
    "name" : "operator",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING"
    },
    "enumValues" : [ "BETWEEN", "NOT_BETWEEN", "EQUAL", "NOT_EQUAL", "GREATER_THAN", "LESS_THAN", "GREATER_OR_EQUAL", "LESS_OR_EQUAL" ]
  }, {
    "name" : "formula1",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  }, {
    "name" : "formula2",
    "requirement" : "OPTIONAL",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## CUSTOM_FORMULA

Allow values that satisfy a custom formula.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| formula | REQUIRED | {"type":"SCALAR","scalarType":"STRING","constraints":[{"type":"NON_BLANK"}]} |  |

```json
{
  "id" : "CUSTOM_FORMULA",
  "summary" : "Allow values that satisfy a custom formula.",
  "fields" : [ {
    "name" : "formula",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "SCALAR",
      "scalarType" : "STRING",
      "constraints" : [ {
        "type" : "NON_BLANK"
      } ]
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
