# GridGrind

GridGrind is a Java 27 / Gradle engine for structured .xlsx automation. Author a canonical
workbook plan through JSON or the pure Java API, provide explicit host authority, and receive
factual results, assertions and execution evidence.

## Start

Build with a Java 27 JDK selected by JAVA_HOME and PATH:

    ./gradlew :cli:shadowJar
    java -jar cli/build/libs/gridgrind.jar help

[Quick start](docs/QUICK_START.md) explains recipes and grants. The CLI uses explicit run,
doctor, catalog, schema, recipe, template, help, version and license commands.

A minimal in-memory plan creates a worksheet without publishing a workbook:

```json
{
  "protocolVersion": 1,
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "NONE"
  },
  "steps": [
    {
      "stepId": "ensure-budget",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Budget"
      },
      "action": {
        "type": "ENSURE_SHEET"
      }
    }
  ]
}
```

Run requires an explicit grant even when publication is NONE. A grant template authorizes
no operations; edit it to the intended resources, operations, targets and publication.

## Contracts

Core owns the canonical model and typed identifiers. JSON owns tolerant intake and derived
schemas; engine owns execution and native workbook adapters. CLI owns transport and workspace IO.
Optional fields are absent on the wire. Protocol version is the integer 1, and wire variant
identity uses type or status.

- [CLI](docs/CLI.md)
- [Execution, authority and publication](docs/EXECUTION_MODEL.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Java API](docs/JAVA_API.md)
- [Generated protocol reference](docs/reference/README.md)
- [POI parity and scope](docs/POI_PARITY.md)
- [Limits](docs/LIMITATIONS.md) and [problems](docs/ERRORS.md)
- [Development](docs/DEVELOPER.md), [fuzzing](docs/FUZZING.md), [releasing](docs/RELEASING.md)

GridGrind's code is under [MPL 2.0](LICENSE). [NOTICE](NOTICE) and the retained component
license texts describe bundled third-party material; [PATENTS.md](PATENTS.md) describes patent
notices. The JAR and launcher archives carry these materials. The license command emits the
consolidated legal surface.
