# Command-line interface

GridGrind accepts an explicit command. Running it without arguments displays overview help.
Only the commands and options below are accepted; command aliases are not supported.

    gridgrind run --request <path|-> --grant <path> [--response <path>] [--execution-root <dir>] [--temp-root <dir>] [--secrets-provider <path>] [--pretty]
    gridgrind doctor --request <path|-> [--grant <path>] [--response <path>] [--execution-root <dir>] [--temp-root <dir>] [--secrets-provider <path>] [--pretty]
    gridgrind catalog [search <text> | show <group>:<id>] [--response <path>] [--pretty]
    gridgrind schema request|grant|response [--response <path>] [--pretty]
    gridgrind recipe list [--response <path>] [--pretty]
    gridgrind recipe show <ID> [--details] [--response <path>] [--pretty]
    gridgrind recipe match <text> [--response <path>] [--pretty]
    gridgrind recipe materialize <ID> --workspace <new-directory> [--pretty]
    gridgrind template request|grant [--response <path>] [--pretty]
    gridgrind help [overview|protocol|guidance] [--format text|structured] [--response <path>] [--pretty]
    gridgrind version [--format text|structured] [--response <path>] [--pretty]
    gridgrind license [--format text|structured] [--response <path>] [--pretty]

Commands and options are separate arguments. Quote search text containing spaces. Recipe IDs
use the uppercase IDs from recipe list. Duplicate, unknown or command-inapplicable options fail.
The format option is restricted to help, version and license; other commands already produce
JSON. Pretty changes JSON indentation. Output goes to stdout unless response names a file.

A request path establishes its containing directory as the execution root. A request read from
stdin uses --request - and requires an explicit --execution-root; that option cannot accompany
a file request. Response files cannot overwrite request files. A run requires an explicit grant;
the grant template is fail-closed and must be edited to authorize the intended paths, targets
and publication. Doctor can perform static diagnosis without a grant.

The grant filename is a host path. Relative `readableResources` FILE paths and
`publicationAuthority` SAVE_AS paths inside that authority document resolve from the process
working directory, not from the grant file or request directory. Use absolute authority paths
or include the workspace prefix. See the [worked path example](QUICK_START.md).

Asset-backed recipes require recipe materialize into a new workspace. The command publishes
the request and declared assets together. It does not execute the workbook plan. Materialization
emits the request on stdout and does not accept a separate response path.

Workspace publication atomically creates a complete directory and rejects a concurrent existing
entry. Supported local Linux/macOS filesystems must provide the qualified no-clobber primitive;
other providers fail before publication. The parent directory and ancestors must remain trusted
and stable during materialization. The executable JAR and launcher scripts enable the native
access required for directory publication. A custom classpath launch needs
`--enable-native-access=ALL-UNNAMED`; a modular launch needs native access for
`dev.erst.gridgrind.cli`. Core authoring and workbook execution do not need this CLI permission.
Local workspace publication requires UTF-8 native filename encoding and losslessly representable
paths. Use a UTF-8 locale; the managed Temurin environment already supplies it. Unqualified
encodings and OS paths that cannot round-trip through UTF-8 fail before publication.

Recipe show emits the executable plan by default. With --details it emits the recipe descriptor,
including portability, intent tags and the exact request profile; it does not execute the plan.

Detected console-write failures, including broken pipes, cause a nonzero process exit.
Transport failure does not undo workbook publication that already completed; inspect the
destination before retrying a run whose response could not be delivered.

The standalone process owns native logging. Stdout or the response file carries one primary
payload. Stderr can contain compact JSON lines: core progress events for a verbose journal,
transport notices (`reason` and `wroteTo`), and native diagnostics with
`type: "NATIVE_DIAGNOSTIC"`, a component (`POI`, `XML`, `SECURITY`, `NATIVE` or `LOGGING`),
a standard severity in `level`, and an optional bounded `exceptionType`. Native diagnostics
contain no log messages, exception messages, stacks or document snippets. They do not determine
the workbook outcome. Native and logging-status events share a limit of 32 per invocation;
one additional `NATIVE_DIAGNOSTIC_LIMIT` frame reports `limit: 32` if more events occur.
Complete stderr frames are serialized through the invocation's diagnostic channel.

The default native logging threshold is WARN. Set `-Dlog4j2.level=INFO`, `DEBUG`, `TRACE` or
`ALL` for more metadata; `LOG4J_LEVEL` is the environment equivalent. JVM settings take
precedence over the environment, and aliases within either source must agree. ERROR, FATAL
and OFF thresholds refuse because they would hide required diagnostics. Other Log4j routing,
provider, status and context settings, classpath logging property files, and unqualified provider
metadata refuse before native initialization with `INVALID_ARGUMENTS`; values are never echoed.
Programmatic `App`, `GridGrindCli` and engine use preserve the embedding application's logging
policy. Their caller owns diagnostic routing.
