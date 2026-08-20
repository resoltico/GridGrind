package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindRequestSurfaceContractText;
import java.util.List;

/** Request and flag-oriented CLI help sections. */
final class GridGrindCliSurfaceRequestSections {
  private GridGrindCliSurfaceRequestSections() {}

  static CliSurface.CliSection request() {
    return new CliSurface.CliSection(
        "Request",
        List.of(
            "protocolVersion is required.",
            "source.type is required. NEW creates a blank workbook."
                + " EXISTING requires source.path.",
            "persistence is required. NONE keeps the workbook in memory only.",
            "planId is optional. When omitted, the response journal omits it.",
            "execution is optional. When omitted, GridGrind uses FULL_XSSF with"
                + " SUMMARY journaling and DO_NOT_CALCULATE without"
                + " markRecalculateOnOpen. When supplied, execution may include any"
                + " subset of execution.mode, execution.journal, execution.calculation, and"
                + " execution.assertionMode. execution.mode is a typed discriminator"
                + " when present; choose type=FULL_XSSF, EVENT_READ, or"
                + " STREAMING_WRITE under the limits above."
                + " Any omitted nested execution field keeps that same default.",
            "execution.journal.level controls journal detail; SUMMARY is the default and"
                + " keeps the response stable by omitting timing telemetry."
                + " VERBOSE records fine-grained phase events in journal.events[].",
            "execution.calculation controls server-side evaluation, cache clearing, and"
                + " open-time recalc flags. strategy.type accepts DO_NOT_CALCULATE,"
                + " EVALUATE_ALL, EVALUATE_TARGETS, and CLEAR_CACHES_ONLY."
                + " EVALUATE_TARGETS also requires strategy.cells[]."
                + " Omit execution.calculation entirely, or supply execution.calculation"
                + " with neither strategy nor markRecalculateOnOpen, to keep the default"
                + " DO_NOT_CALCULATE / false behavior."
                + " markRecalculateOnOpen is otherwise independent.",
            "execution.assertionMode defaults to FAIL_FAST. COLLECT evaluates every assertion"
                + " after the terminal assertion phase begins; no later MUTATION step is legal,"
                + " while INSPECTION steps may still interleave.",
            "Response telemetry is split intentionally: journal.* captures execution-phase"
                + " timing and event chronology, while root persistence/calculation/warnings/"
                + "assertions/inspections carry the authoritative outcome payloads.",
            "NORMAL and VERBOSE journal timings are observational runtime telemetry and vary"
                + " between runs; assert status/category/detail, not literal timings.",
            "Source-backed text and binary fields support INLINE, UTF8_FILE or FILE, and"
                + " STANDARD_INPUT sources.",
            GridGrindRequestSurfaceContractText.standardInputRequiresRequestMessage()
                + " because stdin cannot carry both the request JSON and authored input"
                + " content in one CLI invocation.",
            "formulaEnvironment is optional. When omitted, GridGrind uses the empty"
                + " evaluator environment: externalWorkbooks=[],"
                + " missingWorkbookPolicy=ERROR, udfToolpacks=[]."
                + " When supplied, omitted nested fields keep those same defaults."
                + " missingWorkbookPolicy accepts ERROR or USE_CACHED_VALUE;"
                + " udfToolpacks[] registers named UDF packs for formula evaluation.",
            "persistence.security.encryption writes OOXML AGILE packages only."
                + " encryption.password is required; encryption.cipher defaults to AES_256"
                + " and encryption.hash defaults to SHA_512 when omitted."
                + " Supported ciphers are AES_256 and AES_192; supported hashes are"
                + " SHA_512, SHA_384, and SHA_256."
                + " Legacy STANDARD packages remain readable on inspection but are not"
                + " authorable.",
            "array-formula braces such as {=SUM(A1:A2*B1:B2)} are rejected as"
                + " INVALID_FORMULA.",
            "steps is required. [] represents an empty no-op plan.",
            GridGrindRequestSurfaceContractText.stepKindSummary()
                + " target is a sibling field on each step; it is not nested inside"
                + " action, assertion, or query.",
            "Step order is authoritative. Mutations, assertions, and inspections may be"
                + " interleaved unless the chosen execution mode or calculation strategy"
                + " tightens the ordering contract (for example,"
                + " EVALUATE_ALL/EVALUATE_TARGETS require every MUTATION step to finish"
                + " before observation begins)."));
  }

  static CliSurface.CliDefinitionSection fileWorkflow() {
    return new CliSurface.CliDefinitionSection(
        "File Workflow",
        List.of(
            new CliSurface.DefinitionEntry(
                "No --request flag",
                "read the JSON request from stdin; pass --execution-root so relative"
                    + " request-owned paths resolve from one"
                    + " explicit directory. A bare TTY invocation with no piped request"
                    + " is rejected."),
            new CliSurface.DefinitionEntry(
                "--request <path>",
                "read the JSON request from that file; the request file directory owns"
                    + " request-root resolution."),
            new CliSurface.DefinitionEntry(
                "--execution-root <path>",
                "required when the request JSON arrives on stdin; relative request-owned"
                    + " paths resolve from that directory."),
            new CliSurface.DefinitionEntry(
                "--temp-root <path>",
                "override the scratch-parent path. "
                    + GridGrindRequestSurfaceContractText.cliScratchSpaceSummary()),
            new CliSurface.DefinitionEntry(
                "No --response flag", "write the primary command output to stdout."),
            new CliSurface.DefinitionEntry(
                "--response <path>",
                "write the primary command output to one new file; parent directories are"
                    + " created, but existing files are never replaced implicitly."
                    + " JSON-native payloads stay compact by default; pass --pretty when"
                    + " you want indented JSON."
                    + " Without --response, the command payload is the sole stdout"
                    + " content: CommandError for a rejected command, WorkbookResult"
                    + " for execution, or the command's own discovery, help, or doctor"
                    + " payload. With --response, that payload is written to the new"
                    + " file. When stdout is writable, a response-file write failure"
                    + " recovers the already-rendered payload there unchanged and adds one compact transport notice on stderr;"
                    + " GridGrind never moves a primary payload to stderr."),
            new CliSurface.DefinitionEntry(
                "source.type=EXISTING + source.path", "open an existing workbook from that path."),
            new CliSurface.DefinitionEntry(
                "persistence SAVE_AS.path + ifExists",
                "write to that path with explicit REJECT-or-REPLACE collision policy;"
                    + " parent directories are created."),
            new CliSurface.DefinitionEntry(
                "persistence OVERWRITE", "write back to source.path; no path field is supplied."),
            new CliSurface.DefinitionEntry(
                "Relative CLI flag paths",
                GridGrindRequestSurfaceContractText.cliFlagPathResolutionSummary()),
            new CliSurface.DefinitionEntry(
                "Relative request-owned paths",
                "source.path, persistence paths, source-backed file inputs,"
                    + " formulaEnvironment.externalWorkbooks[*].path, and"
                    + " persistence.security.signature.signature.pkcs12Path follow one rule:"
                    + " "
                    + GridGrindRequestSurfaceContractText.requestOwnedPathResolutionSummary()),
            new CliSurface.DefinitionEntry(
                "Relative FILE hyperlink targets",
                "are analyzed against the persisted workbook path when one exists; use"
                    + " absolute paths for cwd-independent results.")));
  }

  static CliSurface.CliDefinitionSection flags() {
    return new CliSurface.CliDefinitionSection(
        "Flags",
        List.of(
            new CliSurface.DefinitionEntry(
                "--request <path>", "Read the JSON request from a file instead of stdin."),
            new CliSurface.DefinitionEntry(
                "--execution-root <path>",
                "Required when the request JSON arrives on stdin; relative request-owned"
                    + " paths resolve from that directory."),
            new CliSurface.DefinitionEntry(
                "--temp-root <path>",
                "Override the scratch-parent path. "
                    + GridGrindRequestSurfaceContractText.cliScratchSpaceSummary()),
            new CliSurface.DefinitionEntry(
                "--response <path>",
                "Write the primary command output to a new file instead of stdout;"
                    + " GridGrind never replaces an existing response file implicitly."),
            new CliSurface.DefinitionEntry(
                "--format <text|structured>",
                "Render CLI-owned prose surfaces as text or structured JSON. Help,"
                    + " version, and license default to text; JSON-native execution,"
                    + " doctor, request-template, and discovery payloads do not use"
                    + " --format."),
            new CliSurface.DefinitionEntry(
                "--pretty",
                "Indent JSON output. Execution responses, doctor reports, request templates,"
                    + " discovery payloads, and structured help/version/license reports are"
                    + " compact by default and become multi-line JSON when --pretty is"
                    + " supplied."),
            new CliSurface.DefinitionEntry(
                "--doctor-request",
                "Lint one request, preflight source-backed input resolution plus existing"
                    + " workbook-source accessibility, and emit a machine-readable doctor"
                    + " report without mutating a workbook. The doctor response returns"
                    + " warnings plus every independently provable blocking problem it can"
                    + " isolate safely, including structural and constructor-level request"
                    + " intake failures across multiple steps in one pass."),
            new CliSurface.DefinitionEntry(
                "--print-request-template",
                "Print a minimal valid request JSON document with default execution and"
                    + " formula settings omitted."),
            new CliSurface.DefinitionEntry(
                "--print-recipe-catalog",
                "Print the machine-readable unified recipe catalog of built-in examples"
                    + " and CLI-authored task starters. The bare list stays compact;"
                    + " scoped --lookup payloads add richer view-specific detail,"
                    + " including the exact runnable request profile."),
            new CliSurface.DefinitionEntry(
                "--lookup <id>",
                "With --print-recipe, --print-recipe-catalog, or --print-protocol-catalog,"
                    + " print one stable entry by id. With --print-protocol-catalog,"
                    + " that lookup id may also name one top-level category"
                    + " (mutationActionTypes, assertionTypes, inspectionQueryTypes,"
                    + " sourceTypes, persistenceTypes, stepTypes), one nested/plain"
                    + " support group (cellInputTypes, calculationStrategyTypes,"
                    + " executionPolicyInputType, ...), one explicit namespace form"
                    + " (nestedTypes:<group>, plainTypes:<group>), or one qualified"
                    + " top-level entry <category>:<id> such as"
                    + " mutationActionTypes:SET_CELL when ids repeat across groups."
                    + " Scoped lookup payloads may also publish shared top-level notes and"
                    + " entry-local noteRefs when a reusable rule would otherwise be"
                    + " repeated across multiple summaries."),
            new CliSurface.DefinitionEntry(
                "--print-recipe --lookup <id>",
                "Print one built-in example or executable task-starter scenario by id."),
            new CliSurface.DefinitionEntry(
                "--print-recipe-keyword-match --query <text>",
                "Print ranked built-in recipe matches for one English keyword query."
                    + " Use --print-recipe --lookup <id> for the executable starter"
                    + " scenario after you choose a recipe id. After normalization, at least"
                    + " one searchable non-stop-word term must remain."),
            new CliSurface.DefinitionEntry(
                "--print-protocol-catalog",
                "Print the compact protocol-catalog index with group ids, lookup"
                    + " namespace forms, and field-metadata legends such as"
                    + " projectedByFacets, noteRefs, and enumValueDocs."
                    + " Shared reusable notes stay on scoped --lookup payloads instead of"
                    + " bloating the bare index."),
            new CliSurface.DefinitionEntry(
                "--search <text>",
                "With --print-protocol-catalog, perform case-insensitive search across"
                    + " lookup ids, qualified ids, catalog groups, and summaries."
                    + " Search promotes top-level operations first and uses"
                    + " relatedEntryIds on support-group hits."),
            new CliSurface.DefinitionEntry("--help, -h", "Print the short synopsis."),
            new CliSurface.DefinitionEntry(
                "--help-protocol", "Print the authoritative CLI and request grammar only."),
            new CliSurface.DefinitionEntry(
                "--help-guidance",
                "Print workflow guidance, examples, Docker usage, and discovery playbooks."),
            new CliSurface.DefinitionEntry(
                "--version", "Print the GridGrind version and description."),
            new CliSurface.DefinitionEntry(
                "--license", "Print the GridGrind license and third-party notices.")));
  }
}
