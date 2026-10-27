package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.cli.command.CliArguments;
import dev.erst.gridgrind.core.catalog.GridGrindRequestSurfaceContractText;
import java.util.List;

/** Request and flag-oriented CLI help sections. */
final class HelpProtocol {
  private HelpProtocol() {}

  static HelpSurfaceReport.CliSection request() {
    return new HelpSurfaceReport.CliSection(
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
                + " VERBOSE streams compact JSONL progress events to stderr while the primary"
                + " response retains only its execution journal.",
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
            PathGuidance.standardInputRequiresRequestMessage()
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

  static HelpSurfaceReport.CliDefinitionSection fileWorkflow() {
    return new HelpSurfaceReport.CliDefinitionSection(
        "File Workflow",
        List.of(
            new HelpSurfaceReport.DefinitionEntry(
                "No --request flag",
                "read the JSON request from stdin; pass --execution-root so relative"
                    + " request-owned paths resolve from one"
                    + " explicit directory. A bare TTY invocation with no piped request"
                    + " is rejected."),
            new HelpSurfaceReport.DefinitionEntry(
                "--request <path>",
                "read the JSON request from that file; the request file directory owns"
                    + " request-root resolution."),
            new HelpSurfaceReport.DefinitionEntry(
                "--execution-root <path>",
                "required when the request JSON arrives on stdin; relative request-owned"
                    + " paths resolve from that directory."),
            new HelpSurfaceReport.DefinitionEntry(
                "--temp-root <path>",
                "override the scratch-parent path. " + PathGuidance.cliScratchSpaceSummary()),
            new HelpSurfaceReport.DefinitionEntry(
                "No --response flag", "write the primary command output to stdout."),
            new HelpSurfaceReport.DefinitionEntry(
                "Exit status",
                "0 means the command completed successfully; 2 means a CLI argument or"
                    + " request-contract rejection; 1 means an execution, resource, I/O,"
                    + " transport, doctor-finding, or internal failure."),
            new HelpSurfaceReport.DefinitionEntry(
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
            new HelpSurfaceReport.DefinitionEntry(
                "source.type=EXISTING + source.path", "open an existing workbook from that path."),
            new HelpSurfaceReport.DefinitionEntry(
                "persistence SAVE_AS.path + ifExists",
                "write to that path with explicit REJECT-or-REPLACE collision policy;"
                    + " parent directories are created."),
            new HelpSurfaceReport.DefinitionEntry(
                "persistence OVERWRITE", "write back to source.path; no path field is supplied."),
            new HelpSurfaceReport.DefinitionEntry(
                "Relative CLI flag paths", PathGuidance.cliFlagPathResolutionSummary()),
            new HelpSurfaceReport.DefinitionEntry(
                "Relative request-owned paths",
                "source.path, persistence paths, source-backed file inputs,"
                    + " formulaEnvironment.externalWorkbooks[*].path, and"
                    + " persistence.security.signature.signature.pkcs12Path follow one rule:"
                    + " "
                    + PathGuidance.requestOwnedPathResolutionSummary()),
            new HelpSurfaceReport.DefinitionEntry(
                "Relative FILE hyperlink targets",
                "are analyzed against the persisted workbook path when one exists; use"
                    + " absolute paths for cwd-independent results.")));
  }

  static HelpSurfaceReport.CliDefinitionSection flags() {
    return new HelpSurfaceReport.CliDefinitionSection("Options", CliArguments.optionHelp());
  }
}
