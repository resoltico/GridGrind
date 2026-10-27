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

Asset-backed recipes require recipe materialize into a new workspace. The command publishes
the request and declared assets together. It does not execute the workbook plan. Materialization
emits the request on stdout and does not accept a separate response path.

Recipe show emits the executable plan by default. With --details it emits the recipe descriptor,
including portability, intent tags and the exact request profile; it does not execute the plan.
