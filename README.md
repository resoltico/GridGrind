# GridGrind

GridGrind is a Java 27 / Gradle engine for structured .xlsx automation. Author a canonical
workbook plan through JSON or the pure Java API, provide explicit host authority, and receive
factual results, assertions and execution evidence.

## Start

Build with the pinned Eclipse Temurin 27 JDK selected by `JAVA_HOME` and `PATH`. Build,
development and runtime share the same JDK; containers use its Ubuntu 26.04 image.
[Development setup](docs/DEVELOPER.md#toolchain-and-dependencies) describes the exact inputs:

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

## Choose an integration

All Java paths use the pinned JDK. Execution consumes the same canonical plan and explicit
host grant; authoring alone does not execute it.

| Path | Authoring | Execution and outputs | Runtime/components |
| --- | --- | --- | --- |
| JSON with the CLI | Recipes or a request document; schemas and catalog describe the shapes | `doctor` diagnoses; `run` executes and reports facts, assertions and publication evidence | Packaged JAR or launcher with the pinned JDK |
| Java authoring | Typed `GridGrindPlan`, selectors and values | Produces a `WorkbookPlan`; JSON serialization is optional | Core library; add JSON for exchange formats |
| Java embedding | Canonical `WorkbookPlan` and host-owned inputs | Executor or doctor returns the same core outcomes; the host owns streams and logging | Engine and its packaged native provider |
| Container CLI | The same JSON requests, grants and commands | The same CLI contract, with explicitly mounted work and scratch directories | Docker on AMD64 or 64-bit ARM; JDK included |

[FULL_XSSF, EVENT_READ and STREAMING_WRITE](docs/EXECUTION_MODEL.md#execution) have different
capabilities and materialization requirements. A streaming mode does not guarantee bounded
end-to-end heap use. [Raspberry Pi development](docs/DEVELOPER.md#raspberry-pi) uses the same
64-bit ARM contributor and application images.

## Contracts

Core owns the canonical model and typed identifiers. JSON owns tolerant intake and derived
schemas; engine owns execution and native workbook adapters. CLI owns transport and workspace IO.
Absent optional values are omitted from the wire. Protocol version is the integer 1, and wire variant
identity uses type or status.

- [CLI](docs/CLI.md)
- [Execution, authority and publication](docs/EXECUTION_MODEL.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Java API](docs/JAVA_API.md)
- [Generated protocol reference](docs/reference/README.md)
- [POI parity and scope](docs/POI_PARITY.md)
- [Limits](docs/LIMITATIONS.md) and [problems](docs/ERRORS.md)
- [Development](docs/DEVELOPER.md), [fuzzing](docs/FUZZING.md), [releasing](docs/RELEASING.md)

Original GridGrind code and documentation use [AGPL-3.0-only](LICENSE). See
[licensing, source and contribution terms](PROJECT-NOTICE.md) and the
[commercial licensing status](LICENSE-COMMERCIAL.md). [NOTICE](NOTICE) and retained component
license texts describe third-party material; [PATENTS.md](PATENTS.md) describes patent notices.
