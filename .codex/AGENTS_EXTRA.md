# GridGrind Architectural Directives

## Non-Negotiable Policies

**NO EMOJI.** Not in source code, Javadoc, comments, commit messages, documentation, or this
file. Zero tolerance. Remove any emoji encountered.

**NO BACKWARDS COMPATIBILITY.** This is a greenfield project. Hard-break refactors are always
correct. No migration shims, no deprecated bridges, no legacy fallbacks.

---

You are operating within the GridGrind project workspace. This is a strict **Java 26** greenfield
project. Every file you write or modify must embody modern Java 26 idioms precisely. Violations
found in existing code are **primary refactoring targets**, not precedents to follow or work around.

---

## 1. Domain Modeling: Sealed Interfaces and Records

### 1.1 Shape of All Domain Types

Every domain model in GridGrind — commands, requests, responses, contexts, events, errors — is a
`sealed interface` whose only permitted subtypes are `record` implementations. No exceptions.

Forbidden shapes:
- God-records: a single record with optional/nullable fields covering multiple mutually exclusive
  states (e.g., a `status` field that determines which other fields are meaningful).
- Enum-tagged unions: a record or class with an enum discriminator and nullable "padding" fields.
- Class hierarchies with abstract base classes.

Correct shape:
```java
public sealed interface WorkbookSource {
    record New() implements WorkbookSource {}
    record ExistingFile(String path) implements WorkbookSource {}
}
```

### 1.2 Typed Fields and No Shadow Enums

The prohibition on calling `.name()` to produce a wire string is scoped to application-level calls
that cross domain boundaries. It does not apply to Jackson's built-in enum serialization mechanism.
Protocol-owned enums in the `catalog` and `dto` packages (e.g. `FieldRequirement`, `ScalarType`)
are wire-safe when their constant names match the published wire vocabulary by design — but that
match must be verified at the enum definition site, not assumed.

GridGrind example — Apache POI's `CellType.NUMERIC` has `.name()` == `"NUMERIC"`, but the GridGrind
wire contract uses `"NUMBER"`. Use an explicit exhaustive switch:

```java
return switch (cellType) {
    case NUMERIC -> "NUMBER";
    case STRING  -> "STRING";
    case BOOLEAN -> "BOOLEAN";
    case FORMULA -> "FORMULA";
    case BLANK   -> "BLANK";
    case ERROR   -> "ERROR";
    case _STRING -> "STRING";
};
```

### 1.3 Instantiation — No Builders

Instantiate the precise subtype directly:
```java
new WorkbookOperation.SetCell(sheetName, address, value)   // correct
WorkbookOperation.create("SET_CELL", ...)                  // forbidden
```

**Builders are forbidden.** Records have precise constructors — use them. The only permissible
alternative to direct construction is a static factory on the type itself that delegates immediately
to `new` with no additional logic.

---

## 2. Null Discipline

### 2.1 Business Logic — Zero Null Returns

Service methods, domain methods, and all internal logic must not return null. Return an empty
collection, a typed empty-sentinel subtype of a sealed interface, or `Optional` instead.

### 2.2 No Reflection in Tests

Reflection to access or alter private members is unconditionally forbidden in test code. If a
class requires reflection to test, the class has an architectural flaw — fix the architecture.

### 2.3 Protocol Wire Surfaces — No Null Padding

Protocol request, response, and discovery DTOs must encode alternative state in the type system,
not by flattening every possible field onto one interface and returning null when it does not
apply.

Requirements for protocol wire types:
- Do not add `default` accessors whose only purpose is to return null for non-applicable subtypes.
- Model mutually exclusive state with sealed variants, or extract a shared helper record when a
  field is genuinely common.
- External JSON must omit absent properties; it must not serialize explicit `null` placeholders as
  a control protocol for agents or integrations.
- Add or update tests when touching these surfaces so request, response, and discovery JSON all
  stay free of `: null` output.

---

## 3. Immutability and Module Boundaries

### 3.1 JPMS: Module Boundaries Are Architectural

The `module-info.java` descriptors in `excel-foundation`, `engine`, `contract`, `executor`,
`authoring-java`, and `cli` are part of the enforced architecture, not packaging trivia. Preserve
the current product graph from `docs/DEVELOPER_CONTRACT_REPLACEMENT_ADR.md`:
- `authoring-java -> contract -> excel-foundation`
- `cli -> executor -> contract -> excel-foundation`
- `executor -> engine -> excel-foundation`
Do not bypass module boundaries with broadened exports, transitive leaks, or classpath-only
workarounds when the correct fix is to move ownership to the right module.

**When to update `module-info.java`:**
- Adding a new package that Jackson must deserialize: add `opens <pkg> to tools.jackson.databind`.
- Adding a new package whose types are referenced by another module: add `exports <pkg>`.
- Moving a type to a new package: update `exports` and `opens` in the same commit.
- Never add a broad `opens` just to silence a test-time `InaccessibleObjectException`. Diagnose
  which framework needs access and add the narrowly qualified directive instead.

### 3.2 Stream Gatherers — Catalog Construction Only

Use `Stream::gather` only for catalog-construction pipelines that perform encounter-order
deduplication, enrichment from structured metadata, domain validation during emission, or grouping
into a published semantic order. Gatherers are forbidden in executor core flow, engine mutation
logic, or any path where a plain loop is equally clear.

Implementation rules:
- Use `Gatherer.ofSequential(stateFactory, integrator)` for all stateful gatherers. Never use
  the parallel overloads — catalog construction is inherently ordered and sequential.
- Place gatherer factory methods in a dedicated `gather/` sub-package alongside a `CatalogGatherers`
  factory class. Do not inline gatherer lambdas at call sites.
- Name factory methods as verbs: `toOrderedUniqueOrThrow`, `expandFieldsWithMetadata`.

### 3.3 Project-Owned Tooling Seams — Fuzz Driver

The fuzz-driver seam (`FuzzDataProvider` or equivalent) must expose only the operations GridGrind
consumes: `consumeBoolean`, `consumeByte`, `consumeInt`, `consumeRegularDouble`, `remainingBytes`.
A structured fuzz generator must depend on this tiny GridGrind interface rather than a heavyweight
vendor fuzz-data type.

Committed-seed replay must not require native fuzz-driver bootstrap when GridGrind can replay the
same scalar semantics itself.

### 3.4 Long-Running Verification — Semantic Progress

Any local operator gate, replay runner, metadata refresh, or other long-running verification flow
must emit semantic progress pulses and must treat "alive but not advancing" as a failure mode.

Rules:
- A pulse must report completed domain work, not just time passing. Examples: completed test
  class count, completed regression input count, completed packaging stage.
- Track `last progress` separately from `last output`. A chatty log with no new completed work is
  still a stall.
- For staged shell orchestration, emit one stable machine-readable pulse prefix and include the
  current stage, elapsed time, quiet time, stalled time, and latest semantic progress marker.
- When a stage stalls past the configured threshold, capture diagnostics before termination:
  process tree, thread dump for JVMs, open-file or `lsof` snapshot, and recent log tail.
- A plain heartbeat without domain progress is insufficient when the workflow can report
  meaningful units of completion.

---

## 4. Serialization (Jackson 3.x)

### 4.1 Package Names

This project uses Jackson 3.x (`tools.jackson.core:jackson-databind`). Jackson 3.x changed
the group ID and the core/databind implementation packages, but the annotations module kept
its original Java package path. Use the correct package per API surface:

```java
// Annotations module — unchanged in Jackson 3.x
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

// Core/databind APIs — use tools.jackson.* in Jackson 3.x
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.core.JacksonException;
```

### 4.2 Polymorphic Type Registration

Attach `@JsonTypeInfo` and `@JsonSubTypes` to the `sealed interface` declaration, never to
individual record implementations:

```java
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookOperation.SetCell.class,  name = "SET_CELL"),
    @JsonSubTypes.Type(value = WorkbookOperation.SetRange.class, name = "SET_RANGE")
})
public sealed interface WorkbookOperation { ... }
```

- Discriminator property values must be SCREAMING_SNAKE_CASE string literals matching the JSON
  protocol specification exactly.
- Response-side discriminators must echo the corresponding request-side discriminator exactly.
  The `name=` value on a `WorkbookReadResult` subtype must be identical to the `name=` value on
  the corresponding `WorkbookReadOperation` subtype (e.g. `GET_CELLS` on both). Likewise for
  `PersistenceOutcome` and `WorkbookPersistence` (e.g. `NONE`, `SAVE_AS`, `OVERWRITE`). This
  eliminates any client-side translation table between request type and response type.
- Never use `@JsonTypeInfo(use = JsonTypeInfo.Id.CLASS)` — it leaks internal class names into
  the wire format and breaks versioning.

### 4.3 Unknown Field Rejection

The protocol `JsonMapper` must be configured with `FAIL_ON_UNKNOWN_PROPERTIES` enabled. Unknown
fields in a request payload are a protocol violation and must be rejected with
`InvalidRequestShapeException`, not silently ignored. This is already configured in
`GridGrindJson`; do not disable it.

### 4.4 Exception-to-Message Translation: Two-Layer Contract

Jackson exception messages are never safe to surface directly. They contain internal class names,
configuration advice, and source-location metadata that are meaningless or harmful to external
consumers. `GridGrindJson` enforces a two-layer translation boundary:

**Layer 1 — type dispatch (`message` method):** Branch on concrete `JacksonException` subtypes
before touching the message string. Each subtype has a known semantic meaning that maps to a
clean protocol message without requiring any string inspection:
- `InvalidTypeIdException` → `"Unknown type value '...'"` (from `getTypeId()`)
- `UnrecognizedPropertyException` → `"Unknown field '...'"` (from `getPropertyName()`)
- `MismatchedInputException` → delegated to layer 2 (message content varies by subtype)

When Jackson adds a new exception subtype with a known structure, add a new type-dispatch arm
here — never rely on the message string to detect it.

**Layer 2 — message sanitization (`cleanJacksonMessage` + `productOwnedJacksonMessage`):**
For exceptions that reach the message layer, `cleanJacksonMessage` strips all known Jackson
noise patterns (source-location suffixes, subtype descriptions, POJO property references,
configuration advice). `productOwnedJacksonMessage` then pattern-matches the cleaned string
against known message shapes and maps them to canonical protocol messages. The fallback
`return cleaned` at the end is only safe because `cleanJacksonMessage` guarantees the string
is noise-free.

**Maintenance rule:** When a real error from a deployed GridGrind instance produces a message
that leaks Jackson internals, the fix must land in layer 1 (new type arm) or layer 2 (new
stripping rule or new `productOwnedJacksonMessage` pattern). Never add a workaround at the
call site.

---

## 5. Incidental Observation Protocol

When a file read surfaces a defect, a rule violation, or a clear improvement opportunity,
append one entry to `.codex/OBSERVATIONS_INCIDENTAL.txt` and continue the active task.
Do not fix it, do not mention it in chat, do not interrupt the workflow. The file is a
triage backlog reviewed by the project owner — not an action queue.

**Entry format:**
```
------------------------------------------------------------------------
ID: <7-character random alphanumeric, e.g. A3K9PQR>
DATE: YYYY-MM-DD
STATUS: OPEN
FILE: <path>:<line_range>
CATEGORY: DIRECTIVE | DEFECT | COVERAGE | SIMPLIFY | PERF
OBSERVATION: <what is wrong and why it matters>
CURRENT: <minimal code excerpt or pattern name>
FIX: <what change resolves it>
EFFORT: TRIVIAL | MINOR | MODERATE
------------------------------------------------------------------------
```

**ID** is a 7-character random alphanumeric string (uppercase letters and digits). Generate it
randomly — do not use sequential numbers or derive it from the content. It is a stable handle
for referring to an observation across conversations.

**STATUS** lifecycle:
- New entries are always `OPEN`.
- When an observation is fully resolved, update its entry in-place: change `STATUS: OPEN` to
  `STATUS: ACTIONED`. Do not append a duplicate entry.
- Never delete entries. The file is a permanent audit trail.

**DATE** uses the current date from context — no shell call required.

---

## 6. Limitations Registry

All hard constraints enforced by GridGrind — whether derived from Excel structural limits,
Apache POI implementation limits, or GridGrind's own operational constraints — must be
tracked in `docs/LIMITATIONS.md` as a numbered registry. The registry is the single
authoritative source for limit values and their provenance.

### 6.1 Registry Entry Requirements

Each limit must have a stable `LIM-NNN` identifier (three-digit zero-padded integer, assigned
sequentially). Entries must not be renumbered once assigned — if a limit is retired, mark it
as retired rather than reusing the ID.

Every registry entry must state:
- The `LIM-NNN` identifier
- The category (GridGrind operational, Excel structural, Apache POI implementation, or protocol)
- The enforced limit value
- The error code raised when the limit is exceeded
- The exact error message text
- Which operations the limit applies to
- The code reference where the limit is enforced
- The UX reference where the limit is surfaced (catalog summary, `--help` text, or both)
- A brief explanation of why the limit exists and where it originates

### 6.2 Synchronization — Three Required Surfaces

A limit is only fully registered when it is present in all three surfaces simultaneously:

1. **Registry entry** in `docs/LIMITATIONS.md` with the `LIM-NNN` ID.
2. **Enforcement code comment** — the validation site in source code must carry a trailing
   `// LIM-NNN` comment on the constant declaration or the validation method.
3. **UX string** — the protocol catalog summary and/or the `--help` `Limits:` section must
   state the limit value so agents and users can discover it before constructing a request.

A limit that exists in enforcement code but is absent from the registry, or present in the
registry but absent from the UX strings, is incomplete. Complete all three surfaces together.

### 6.3 When a Limit Changes

When a limit value changes, update all three surfaces atomically in the same commit:
- Update the registry entry value in `docs/LIMITATIONS.md`.
- Update the constant or validation in source code.
- Update the catalog summary and/or `--help` text.
- Add a `CHANGELOG.md` entry under `[Unreleased]` describing the change.

Never update one surface without the others. A stale UX string is a user-facing defect.

### 6.4 Adding a New Limit

1. Assign the next `LIM-NNN` ID from the registry (check the highest existing ID and
   increment by one).
2. Add the enforcement code comment `// LIM-NNN` at the validation site.
3. Add the registry entry to `docs/LIMITATIONS.md`.
4. Add or update the relevant catalog summary and `--help` `Limits:` section.
5. Add a `CHANGELOG.md` entry.
