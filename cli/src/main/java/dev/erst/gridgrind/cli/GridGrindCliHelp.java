package dev.erst.gridgrind.cli;

import static dev.erst.gridgrind.cli.GridGrindCliHelpRenderSupport.formatRecipes;
import static dev.erst.gridgrind.cli.GridGrindCliHelpRenderSupport.renderCommandExample;
import static dev.erst.gridgrind.cli.GridGrindCliHelpRenderSupport.renderCommandSection;
import static dev.erst.gridgrind.cli.GridGrindCliHelpRenderSupport.renderCoordinateSystems;
import static dev.erst.gridgrind.cli.GridGrindCliHelpRenderSupport.renderDefinitions;
import static dev.erst.gridgrind.cli.GridGrindCliHelpRenderSupport.renderDiscovery;
import static dev.erst.gridgrind.cli.GridGrindCliHelpRenderSupport.renderReferences;
import static dev.erst.gridgrind.cli.GridGrindCliHelpRenderSupport.renderSection;
import static dev.erst.gridgrind.cli.GridGrindCliHelpRenderSupport.renderWorkflows;

import dev.erst.gridgrind.cli.discovery.GridGrindRecipeCatalog;
import java.util.List;
import java.util.Objects;

/** Renders the public CLI help text from CLI-owned discovery metadata. */
public final class GridGrindCliHelp {
  private GridGrindCliHelp() {}

  /** Renders one named CLI help surface for one packaged product identity. */
  public static String helpText(
      CliCommand.HelpTopic topic,
      String version,
      String description,
      String documentRef,
      String containerTag) {
    Objects.requireNonNull(topic, "topic must not be null");
    Objects.requireNonNull(version, "version must not be null");
    Objects.requireNonNull(description, "description must not be null");
    Objects.requireNonNull(documentRef, "documentRef must not be null");
    Objects.requireNonNull(containerTag, "containerTag must not be null");

    CliSurface cliSurface = GridGrindProtocolCatalogCliSurface.CLI_SURFACE;
    return switch (topic) {
      case OVERVIEW ->
          String.join(
                  "\n\n",
                  productHeader(version, description),
                  renderOverview(cliSurface, documentRef))
              + "\n";
      case PROTOCOL ->
          String.join("\n\n", productHeader(version, description), renderProtocolHelp(cliSurface))
              + "\n";
      case GUIDANCE ->
          String.join(
                  "\n\n",
                  productHeader(version, description),
                  renderGuidanceHelp(cliSurface, containerTag),
                  renderReferences(cliSurface.docs(), documentRef))
              + "\n";
    };
  }

  /** Returns the shared two-line product header used by help and version surfaces. */
  public static String productHeader(String version, String description) {
    Objects.requireNonNull(version, "version must not be null");
    Objects.requireNonNull(description, "description must not be null");
    return "GridGrind " + version + "\n" + description;
  }

  private static String renderOverview(CliSurface cliSurface, String documentRef) {
    return String.join(
        "\n\n",
        renderCommandSection(cliSurface.usage()),
        renderDefinitions(
            new CliSurface.CliDefinitionSection(
                "Primary Commands",
                List.of(
                    new CliSurface.DefinitionEntry(
                        "--request <path>",
                        "Execute one request file and emit a structured execution response."),
                    new CliSurface.DefinitionEntry(
                        "--execution-root <path>",
                        "Execute or doctor one stdin-backed request from one explicit request"
                            + " root."),
                    new CliSurface.DefinitionEntry(
                        "--doctor-request",
                        "Lint one request and emit a structured doctor report without workbook"
                            + " mutation."),
                    new CliSurface.DefinitionEntry(
                        "--print-request-template",
                        "Emit the canonical minimal request JSON skeleton."),
                    new CliSurface.DefinitionEntry(
                        "--print-recipe --lookup <id>",
                        "Emit one built-in example or task-starter request by stable id."),
                    new CliSurface.DefinitionEntry(
                        "--print-recipe-catalog [--lookup <id>]",
                        "List every built-in recipe or print one recipe descriptor by id."),
                    new CliSurface.DefinitionEntry(
                        "--print-recipe-keyword-match --query <text>",
                        "Rank likely built-in recipe ids for one natural-language query."),
                    new CliSurface.DefinitionEntry(
                        "--print-protocol-catalog",
                        "Emit the compact authoritative protocol-catalog index."),
                    new CliSurface.DefinitionEntry(
                        "--print-protocol-catalog --lookup <lookup-id>",
                        "Emit one authoritative catalog entry or one scoped type-group/category"
                            + " payload by stable lookup id."),
                    new CliSurface.DefinitionEntry(
                        "--print-protocol-catalog --search <text>",
                        "Search authoritative catalog ids and summaries; rerun --lookup for"
                            + " one full entry or type-group definition."),
                    new CliSurface.DefinitionEntry("--help, -h", "Show the short synopsis."),
                    new CliSurface.DefinitionEntry(
                        "--help-protocol", "Show the authoritative CLI and request grammar only."),
                    new CliSurface.DefinitionEntry(
                        "--help-guidance",
                        "Show workflows, examples, Docker usage, and discovery guidance."),
                    new CliSurface.DefinitionEntry(
                        "--version", "Print the packaged GridGrind version and description."),
                    new CliSurface.DefinitionEntry(
                        "--license", "Print the packaged license and third-party notices.")))),
        renderSection(
            new CliSurface.CliSection(
                "Command Rules",
                List.of(
                    "Every invocation accepts exactly one primary command.",
                    "A bare gridgrind invocation expects one request JSON document on standard"
                        + " input together with --execution-root <path>, or one --request"
                        + " <path>.",
                    "Without --response, the command payload is the sole stdout content. With"
                        + " --response, GridGrind writes that payload to one new file; only a"
                        + " response-file write failure emits one compact transport notice on"
                        + " stderr after recovering the payload on stdout.",
                    "Use --format structured when you want JSON help, version, or license"
                        + " discovery instead of prose.",
                    "Use --pretty when you want indented JSON instead of the compact default"
                        + " on execution, doctor, request-template, discovery, or structured"
                        + " help/version/license output."))),
        renderSection(
            new CliSurface.CliSection(
                "Next Commands",
                List.of(
                    "--help is the short synopsis.",
                    "--help-protocol is the authoritative CLI and request grammar.",
                    "--help-guidance is the workflow, discovery, and example playbook.",
                    "The docs index lives under --help-guidance and repository docs such as "
                        + documentRef
                        + "/docs/QUICK_REFERENCE.md"))));
  }

  private static String renderProtocolHelp(CliSurface cliSurface) {
    return String.join(
        "\n\n",
        renderCommandSection(cliSurface.usage()),
        renderDefinitions(cliSurface.flags()),
        renderSection(
            new CliSurface.CliSection(
                "Authoritative Contract Scope",
                List.of(
                    "This help surface describes the authoritative CLI and request contract"
                        + " only.",
                    "The bare --print-protocol-catalog command emits only the compact"
                        + " index. Scoped --lookup payloads may annotate conditional fields"
                        + " with projectedByFacets, noteRefs, and enumValueDocs, and may"
                        + " publish shared notes when one stable rule applies across"
                        + " multiple entries.",
                    "Workflow playbooks, examples, Docker usage, and catalog walk-throughs live"
                        + " under --help-guidance."))),
        renderSection(cliSurface.execution()),
        renderDefinitions(cliSurface.limits()),
        renderSection(cliSurface.request()),
        renderDefinitions(cliSurface.fileWorkflow()),
        renderCoordinateSystems(cliSurface.coordinateSystems()),
        renderSection(
            new CliSurface.CliSection(
                "Request Shape Facts",
                List.of(
                    "execution.mode is one typed variant; choose type=FULL_XSSF, EVENT_READ, or"
                        + " STREAMING_WRITE.",
                    "formulaEnvironment.missingWorkbookPolicy accepts ERROR or"
                        + " USE_CACHED_VALUE.",
                    "formulaEnvironment.udfToolpacks[] registers named template-backed UDF packs"
                        + " for formula evaluation.",
                    "EVALUATE_TARGETS requires strategy.cells[] and each target must identify an"
                        + " existing formula cell.",
                    "stepId must be unique within steps[] and must match [A-Za-z0-9._-]+.",
                    "VERBOSE stderr emits one line per phase event as timestamp CATEGORY detail"
                        + " with optional stepIndex/stepId pairs."))));
  }

  private static String renderGuidanceHelp(CliSurface cliSurface, String containerTag) {
    String recipeCatalogEntries = formatRecipes(GridGrindRecipeCatalog.catalog().recipes());
    return String.join(
        "\n\n",
        renderSection(
            new CliSurface.CliSection(
                "Operator Guidance Scope",
                List.of(
                    "This help surface contains workflows, examples, and operational playbooks.",
                    "It does not extend the request grammar; use --help-protocol for the"
                        + " authoritative contract."))),
        renderSection(
            new CliSurface.CliSection(
                "Quick Start",
                List.of(
                    "1. Print a minimal request: gridgrind --print-request-template --response"
                        + " request.json",
                    "2. Print a task starter: gridgrind --print-recipe --lookup DASHBOARD"
                        + " --response task-request.json",
                    "3. Preflight the starter: gridgrind --doctor-request --request"
                        + " task-request.json --response doctor.json",
                    "4. Execute the request: gridgrind --request task-request.json --response"
                        + " response.json",
                    "5. For stdin-driven runs, pass one explicit request root:"
                        + " gridgrind --execution-root . < request.json"))),
        renderWorkflows(
            new CliSurface.CliWorkflowSection(
                "Workflow Playbooks", cliSurface.workflows().entries())),
        renderCommandExample(cliSurface.stdinExample(), containerTag),
        renderCommandExample(cliSurface.dockerFileExample(), containerTag),
        renderDiscovery(cliSurface.discovery(), recipeCatalogEntries));
  }
}
