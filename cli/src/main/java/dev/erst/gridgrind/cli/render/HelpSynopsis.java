package dev.erst.gridgrind.cli.render;

import java.util.List;

/** Synopsis-first CLI help sections. */
final class HelpSynopsis {
  private HelpSynopsis() {}

  static HelpSurfaceReport.CliSection usage() {
    return new HelpSurfaceReport.CliSection(
        "Usage",
        List.of(
            "gridgrind run --request <path|-> --grant <path> [--response <path>] [--pretty]",
            "gridgrind doctor --request <path|-> [--grant <path>] [--response <path>] [--pretty]",
            "gridgrind catalog [search <text> | show <group>:<id>] [--response <path>] [--pretty]",
            "gridgrind schema request|grant|response [--response <path>] [--pretty]",
            "gridgrind recipe list | show <ID> | match <text> [--response <path>] [--pretty]",
            "gridgrind recipe materialize <ID> --workspace <new-directory> [--pretty]",
            "gridgrind template request|grant [--response <path>] [--pretty]",
            "gridgrind help [overview|protocol|guidance] [--format text|structured]",
            "gridgrind version | license [--format text|structured]",
            "For stdin requests use --request - with --execution-root <directory>.",
            "Run and doctor also accept --temp-root and --secrets-provider paths."));
  }

  static HelpSurfaceReport.CliWorkflowSection workflows() {
    return new HelpSurfaceReport.CliWorkflowSection(
        "Workflows",
        List.of(
            new HelpSurfaceReport.WorkflowEntry(
                "Discover",
                List.of(
                    "List recipes: gridgrind recipe list",
                    "Match a goal: gridgrind recipe match \"monthly dashboard\"",
                    "Find protocol capabilities: gridgrind catalog search \"chart title\"",
                    "Inspect an exact shape: gridgrind catalog show mutationActionTypes:SET_CELL")),
            new HelpSurfaceReport.WorkflowEntry(
                "Prepare and authorize",
                List.of(
                    "Create a request: gridgrind recipe show BUDGET --response budget-request.json",
                    "Create fail-closed authority: gridgrind template grant --response grant.json",
                    "Edit the grant to authorize the request's exact resources, targets and publication.",
                    "Diagnose: gridgrind doctor --request budget-request.json --grant grant.json")),
            new HelpSurfaceReport.WorkflowEntry(
                "Execute",
                List.of(
                    "Run: gridgrind run --request budget-request.json --grant grant.json --response response.json",
                    "Materialize asset-backed recipes into a new workspace before execution.",
                    "Inspect the structured result, warnings, assertions and publication outcome."))));
  }

  static HelpSurfaceReport.CliSection execution() {
    return new HelpSurfaceReport.CliSection(
        "Execution",
        List.of(
            "Without --response, the command payload is the sole stdout content. With --response,"
                + " that payload is written to a new file; existing response files are not replaced.",
            "A response-file write failure with writable stdout recovers the already-rendered"
                + " payload there unchanged and adds one compact transport notice on stderr."
                + " GridGrind never moves a primary payload to stderr.",
            "Use --format structured for JSON help, version and license reports. Use --pretty"
                + " to indent JSON output.",
            "GridGrind executes ordered steps in sequence, then saves the workbook (unless"
                + " persistence is NONE); if any step fails, no workbook is written.",
            "source.type=NEW starts with zero sheets; ENSURE_SHEET creates the first" + " sheet.",
            "execution is optional; omit it for FULL_XSSF / SUMMARY / DO_NOT_CALCULATE,"
                + " or supply only the nested execution.mode, execution.journal, and"
                + " execution.calculation fields that need non-default behavior."));
  }

  static HelpSurfaceReport.CliDefinitionSection limits() {
    return new HelpSurfaceReport.CliDefinitionSection(
        "Limits",
        dev.erst.gridgrind.core.limits.Limit.all().stream()
            .map(
                limit ->
                    new HelpSurfaceReport.DefinitionEntry(
                        limit.id() + " " + limit.title(), limit.value()))
            .toList());
  }
}
