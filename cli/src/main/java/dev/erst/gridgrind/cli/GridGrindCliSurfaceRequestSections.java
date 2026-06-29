package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
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
            "execution is required. execution.mode is a typed discriminator;"
                + " choose type=FULL_XSSF, EVENT_READ, or STREAMING_WRITE under the"
                + " limits above.",
            "execution.journal.level controls journal detail; SUMMARY is the default and"
                + " keeps the response stable by omitting timing telemetry."
                + " VERBOSE also streams live phase events to stderr as timestamped CATEGORY"
                + " detail lines with optional stepIndex/stepId pairs.",
            "execution.calculation controls server-side evaluation, cache clearing, and"
                + " open-time recalc flags. strategy.type accepts DO_NOT_CALCULATE,"
                + " EVALUATE_ALL, EVALUATE_TARGETS, and CLEAR_CACHES_ONLY."
                + " EVALUATE_TARGETS also requires strategy.cells[]."
                + " markRecalculateOnOpen is independent.",
            "Response telemetry is split intentionally: journal.* captures execution-phase"
                + " timing and event chronology, while root calculation/warnings/assertions/"
                + "inspections carry the authoritative outcome payloads.",
            "NORMAL and VERBOSE journal timings are observational runtime telemetry and vary"
                + " between runs; assert status/category/detail, not literal timings.",
            "Source-backed text and binary fields support INLINE, UTF8_FILE or FILE, and"
                + " STANDARD_INPUT sources.",
            GridGrindContractText.standardInputRequiresRequestMessage()
                + " because stdin cannot carry both the request JSON and authored input"
                + " content in one CLI invocation.",
            "formulaEnvironment is required. The empty-environment shape is"
                + " externalWorkbooks=[], missingWorkbookPolicy=ERROR, udfToolpacks=[]."
                + " missingWorkbookPolicy accepts ERROR or USE_CACHED_VALUE;"
                + " udfToolpacks[] registers named UDF packs for formula evaluation.",
            "array-formula braces such as {=SUM(A1:A2*B1:B2)} are rejected as"
                + " INVALID_FORMULA.",
            "steps is required. [] represents an empty no-op plan.",
            GridGrindContractText.stepKindSummary()
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
                    + " request-owned paths and execution temp files resolve from one"
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
                "override execution scratch space. Without it, temp files resolve under"
                    + " .gridgrind/tmp inside the request root or explicit"
                    + " --execution-root."),
            new CliSurface.DefinitionEntry(
                "No --response flag", "write the primary command output to stdout."),
            new CliSurface.DefinitionEntry(
                "--response <path>",
                "write the primary command output to one new file; parent directories are"
                    + " created, but existing files are never replaced implicitly."
                    + " Without --response, CLI argument errors and request-content failure"
                    + " reports stay on stderr, while executed GridGrindResponse payloads"
                    + " stay on stdout even when status=FAILED."
                    + " Execution writes the JSON response, doctoring writes the doctor"
                    + " report, and help or discovery commands write their rendered text"
                    + " or JSON payload. Non-success results also emit one stderr pointer"
                    + " line naming the file: CLI argument errors write 'CLI failure"
                    + " report', request-content errors write 'request failure report',"
                    + " execution failures write 'response', and doctor failures write"
                    + " 'doctor report'."),
            new CliSurface.DefinitionEntry(
                "source.type=EXISTING + source.path", "open an existing workbook from that path."),
            new CliSurface.DefinitionEntry(
                "persistence SAVE_AS.path",
                "write a new workbook to that path; parent directories are created."),
            new CliSurface.DefinitionEntry(
                "persistence OVERWRITE", "write back to source.path; no path field is supplied."),
            new CliSurface.DefinitionEntry(
                "Relative CLI flag paths", GridGrindContractText.cliFlagPathResolutionSummary()),
            new CliSurface.DefinitionEntry(
                "Relative request-owned paths",
                "source.path, persistence paths, source-backed file inputs,"
                    + " formulaEnvironment.externalWorkbooks[*].path, and"
                    + " persistence.security.signature.pkcs12Path follow one rule:"
                    + " "
                    + GridGrindContractText.requestOwnedPathResolutionSummary()),
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
                    + " paths and execution temp files resolve from that directory."),
            new CliSurface.DefinitionEntry(
                "--temp-root <path>",
                "Override execution scratch space. Without it, GridGrind uses"
                    + " .gridgrind/tmp inside the request root or explicit"
                    + " --execution-root."),
            new CliSurface.DefinitionEntry(
                "--response <path>",
                "Write the primary command output to a new file instead of stdout;"
                    + " GridGrind never replaces an existing response file implicitly."),
            new CliSurface.DefinitionEntry(
                "--format <text|structured>",
                "Render CLI-owned prose surfaces as text or structured JSON. Help,"
                    + " version, and license default to text; JSON-native execution,"
                    + " doctor, and discovery payloads remain JSON in either mode."),
            new CliSurface.DefinitionEntry(
                "--doctor-request",
                "Lint one request, preflight source-backed input resolution plus existing"
                    + " workbook-source accessibility, and emit a machine-readable doctor"
                    + " report without mutating a workbook. The doctor response returns"
                    + " warnings plus every independently provable blocking problem."),
            new CliSurface.DefinitionEntry(
                "--print-request-template", "Print a minimal valid request JSON document."),
            new CliSurface.DefinitionEntry(
                "--print-example-catalog",
                "Print the machine-readable built-in example catalog, including the stable"
                    + " workspaceMode portability contract plus any"
                    + " requiredWorkspacePaths for asset-backed example ids."),
            new CliSurface.DefinitionEntry(
                "--print-task-catalog",
                "Print the machine-readable task catalog of high-level office-work"
                    + " recipes, including starter.requestFileName,"
                    + " starter.workspaceMode, and starter.requiredWorkspacePaths."),
            new CliSurface.DefinitionEntry(
                "--lookup <id>",
                "With --print-example, --print-task-catalog, --print-task-plan, or"
                    + " --print-protocol-catalog, print one stable entry by id (SET_CELL,"
                    + " ENSURE_SHEET, …), one nested/plain type-group descriptor by group"
                    + " name (cellInputTypes, calculationStrategyTypes, …), or one"
                    + " top-level category array by name (mutationActionTypes,"
                    + " assertionTypes, inspectionQueryTypes, sourceTypes,"
                    + " persistenceTypes, stepTypes). Qualify as <category>:<id>"
                    + " (e.g. mutationActionTypes:SET_CELL) when ids repeat across groups."),
            new CliSurface.DefinitionEntry(
                "--print-task-plan --lookup <id>",
                "Print one executable starter scenario for one task id."),
            new CliSurface.DefinitionEntry(
                "--print-task-keyword-match --query <text>",
                "Print ranked CLI-owned task matches for one English keyword query."
                    + " Use --print-task-plan --lookup <id> for the executable starter"
                    + " scenario after you choose a task id. After normalization, at least"
                    + " one searchable non-stop-word term must remain."),
            new CliSurface.DefinitionEntry(
                "--print-protocol-catalog",
                "Print the compact protocol-catalog index with group ids and lookup"
                    + " namespace forms."),
            new CliSurface.DefinitionEntry(
                "--full",
                "With --print-protocol-catalog, print the complete machine-readable"
                    + " protocol catalog."),
            new CliSurface.DefinitionEntry(
                "--search <text>",
                "With --print-protocol-catalog, perform case-insensitive search across"
                    + " lookup ids, qualified ids, catalog groups, and summaries."
                    + " Search promotes top-level operations first and uses"
                    + " relatedEntryIds on support-group hits."),
            new CliSurface.DefinitionEntry(
                "--print-example --lookup <id>", "Print one built-in generated example request."),
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
