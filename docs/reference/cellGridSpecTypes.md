<!-- Generated from core metadata; do not edit. -->
# cellGridSpecTypes

## TYPED

Write one rectangular grid expressed as per-cell typed rows.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"LIST","elementShape":{"type":"NESTED_TYPE_GROUP","group":"cellValueSpecTypes"}}} |  |

```json
{
  "id" : "TYPED",
  "summary" : "Write one rectangular grid expressed as per-cell typed rows.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "LIST",
        "elementShape" : {
          "type" : "NESTED_TYPE_GROUP",
          "group" : "cellValueSpecTypes"
        }
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## TEXT

Write one rectangular grid using homogeneous string rows.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}}} |  |

```json
{
  "id" : "TEXT",
  "summary" : "Write one rectangular grid using homogeneous string rows.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "LIST",
        "elementShape" : {
          "type" : "SCALAR",
          "scalarType" : "STRING"
        }
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## NUMBER

Write one rectangular grid using homogeneous numeric rows.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"NUMBER"}}} |  |

```json
{
  "id" : "NUMBER",
  "summary" : "Write one rectangular grid using homogeneous numeric rows.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "LIST",
        "elementShape" : {
          "type" : "SCALAR",
          "scalarType" : "NUMBER"
        }
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## BOOLEAN

Write one rectangular grid using homogeneous boolean rows.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"BOOLEAN"}}} |  |

```json
{
  "id" : "BOOLEAN",
  "summary" : "Write one rectangular grid using homogeneous boolean rows.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "LIST",
        "elementShape" : {
          "type" : "SCALAR",
          "scalarType" : "BOOLEAN"
        }
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## ERROR

Write one rectangular grid using homogeneous stored Excel error literal rows.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}}} |  |

```json
{
  "id" : "ERROR",
  "summary" : "Write one rectangular grid using homogeneous stored Excel error literal rows.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "LIST",
        "elementShape" : {
          "type" : "SCALAR",
          "scalarType" : "STRING"
        }
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

Write one rectangular grid using homogeneous date rows.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}}} |  |

```json
{
  "id" : "DATE",
  "summary" : "Write one rectangular grid using homogeneous date rows.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "LIST",
        "elementShape" : {
          "type" : "SCALAR",
          "scalarType" : "STRING"
        }
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## DATE_TIME

Write one rectangular grid using homogeneous date-time rows.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}}} |  |

```json
{
  "id" : "DATE_TIME",
  "summary" : "Write one rectangular grid using homogeneous date-time rows.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "LIST",
        "elementShape" : {
          "type" : "SCALAR",
          "scalarType" : "STRING"
        }
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```

## FORMULA

Write one rectangular grid using homogeneous formula text rows.

| Field | Requirement | Shape | Default |
| --- | --- | --- | --- |
| cells | REQUIRED | {"type":"LIST","elementShape":{"type":"LIST","elementShape":{"type":"SCALAR","scalarType":"STRING"}}} |  |

```json
{
  "id" : "FORMULA",
  "summary" : "Write one rectangular grid using homogeneous formula text rows.",
  "fields" : [ {
    "name" : "cells",
    "requirement" : "REQUIRED",
    "shape" : {
      "type" : "LIST",
      "elementShape" : {
        "type" : "LIST",
        "elementShape" : {
          "type" : "SCALAR",
          "scalarType" : "STRING"
        }
      }
    },
    "enumValues" : [ ]
  } ],
  "targetSelectors" : [ ]
}
```
