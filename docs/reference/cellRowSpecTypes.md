<!-- Generated from core metadata; do not edit. -->
# cellRowSpecTypes

## TYPED

Append one row expressed as a per-cell typed value list.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"cellValueSpecTypes"}} |  |

```json
{
  "id" : "TYPED",
  "summary" : "Append one row expressed as a per-cell typed value list.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "NESTED_TYPE_GROUP",
        "group" : "cellValueSpecTypes"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## TEXT

Append one row using a compact homogeneous string list.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "TEXT",
  "summary" : "Append one row using a compact homogeneous string list.",
  "fields" : [ {
    "name" : "cells",
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

## NUMBER

Append one row using a compact homogeneous numeric list.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"NUMBER"}} |  |

```json
{
  "id" : "NUMBER",
  "summary" : "Append one row using a compact homogeneous numeric list.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "NUMBER"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## BOOLEAN

Append one row using a compact homogeneous boolean list.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"BOOLEAN"}} |  |

```json
{
  "id" : "BOOLEAN",
  "summary" : "Append one row using a compact homogeneous boolean list.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "BOOLEAN"
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## ERROR

Append one row using a compact homogeneous stored Excel error literal list.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "ERROR",
  "summary" : "Append one row using a compact homogeneous stored Excel error literal list.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "SCALAR",
        "scalarType" : "STRING"
      }
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

Append one row using a compact homogeneous date list.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "DATE",
  "summary" : "Append one row using a compact homogeneous date list.",
  "fields" : [ {
    "name" : "cells",
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

## DATE_TIME

Append one row using a compact homogeneous date-time list.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "DATE_TIME",
  "summary" : "Append one row using a compact homogeneous date-time list.",
  "fields" : [ {
    "name" : "cells",
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

## FORMULA

Append one row using a compact homogeneous formula text list.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}} |  |

```json
{
  "id" : "FORMULA",
  "summary" : "Append one row using a compact homogeneous formula text list.",
  "fields" : [ {
    "name" : "cells",
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
