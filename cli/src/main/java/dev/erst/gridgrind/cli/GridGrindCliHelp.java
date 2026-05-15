package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.ShippedExampleEntry;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import java.io.IOException;
import java.util.List;
import java.util.Objects;
import java.util.regex.Pattern;

/** Renders the public CLI help text from CLI-owned discovery metadata. */
public final class GridGrindCliHelp {
  private static final int HELP_TEXT_WIDTH = 96;
  private static final Pattern WHITESPACE_PATTERN = Pattern.compile("\\s+");

  private GridGrindCliHelp() {}

  /** Supplies request-template bytes so tests can cover success and failure rendering paths. */
  @FunctionalInterface
  interface RequestTemplateBytesSupplier {
    /** Returns one UTF-8 request-template byte sequence. */
    byte[] get() throws IOException;
  }

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
        renderSection(
            new CliSurface.CliSection(
                "Usage",
                List.of(
                    "gridgrind --request <path> [--response <path>]",
                    "gridgrind [--response <path>] < request.json",
                    "gridgrind --doctor-request [--request <path>] [--response <path>]",
                    "gridgrind --help-protocol | --help-guidance [--response <path>]"))),
        renderDefinitions(
            new CliSurface.CliDefinitionSection(
                "Primary Commands",
                List.of(
                    new CliSurface.DefinitionEntry(
                        "--request <path>",
                        "Execute one request file and emit a structured execution response."),
                    new CliSurface.DefinitionEntry(
                        "--doctor-request",
                        "Lint one request and emit a structured doctor report without workbook"
                            + " mutation."),
                    new CliSurface.DefinitionEntry(
                        "--print-request-template",
                        "Emit the canonical minimal request JSON skeleton."),
                    new CliSurface.DefinitionEntry(
                        "--print-example-catalog",
                        "List built-in example ids plus portability and required-path details."),
                    new CliSurface.DefinitionEntry(
                        "--print-example --lookup <id>", "Emit one built-in example request."),
                    new CliSurface.DefinitionEntry(
                        "--print-task-catalog [--lookup <id>]",
                        "List CLI-owned task recipes or print one task entry by id."),
                    new CliSurface.DefinitionEntry(
                        "--print-task-plan --lookup <id>",
                        "Emit one runnable starter request for a stable task id."),
                    new CliSurface.DefinitionEntry(
                        "--print-task-keyword-match --query <text>",
                        "Rank likely task ids for one natural-language query."),
                    new CliSurface.DefinitionEntry(
                        "--print-protocol-catalog",
                        "Emit the authoritative machine-readable protocol catalog."),
                    new CliSurface.DefinitionEntry(
                        "--print-protocol-catalog --lookup <id>|<group>:<id>",
                        "Emit one authoritative catalog entry or one type group by stable lookup"
                            + " id."),
                    new CliSurface.DefinitionEntry(
                        "--print-protocol-catalog --search <text>",
                        "Search authoritative catalog ids, summaries, and related support"
                            + " groups."),
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
                "Quick Start",
                List.of(
                    "1. Emit a request skeleton: gridgrind --print-request-template --response"
                        + " request.json",
                    "2. Pick a starter request: gridgrind --print-task-plan --lookup DASHBOARD"
                        + " --response request.json",
                    "3. Preflight before execution: gridgrind --doctor-request --request"
                        + " request.json --response doctor.json",
                    "4. Execute explicitly: gridgrind --request request.json --response"
                        + " response.json"))),
        renderSection(
            new CliSurface.CliSection(
                "Command Rules",
                List.of(
                    "Every invocation accepts exactly one primary command.",
                    "A bare gridgrind invocation expects one request JSON document on standard"
                        + " input or --request <path>.",
                    "Failure payloads for CLI argument and lookup errors are emitted as compact"
                        + " CLI failure reports instead of execution journals.",
                    "--help is the short synopsis. Use --help-protocol for the contract and"
                        + " --help-guidance for workflows."))),
        renderReferences(
            new CliSurface.CliReferenceSection("Docs At A Glance", cliSurface.docs().entries()),
            documentRef));
  }

  private static String renderProtocolHelp(CliSurface cliSurface) {
    return String.join(
        "\n\n",
        renderSection(cliSurface.usage()),
        renderDefinitions(cliSurface.flags()),
        renderSection(
            new CliSurface.CliSection(
                "Authoritative Contract Scope",
                List.of(
                    "This help surface describes the authoritative CLI and request contract"
                        + " only.",
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
    String discoveryExamples = formatExamples(GridGrindShippedExamples.catalog().examples());
    return String.join(
        "\n\n",
        renderSection(
            new CliSurface.CliSection(
                "Operator Guidance Scope",
                List.of(
                    "This help surface contains workflows, examples, and operational playbooks.",
                    "It does not extend the request grammar; use --help-protocol for the"
                        + " authoritative contract."))),
        renderWorkflows(cliSurface.workflows()),
        renderCommandExample(cliSurface.stdinExample(), containerTag),
        renderCommandExample(cliSurface.dockerFileExample(), containerTag),
        renderDiscovery(cliSurface.discovery(), discoveryExamples));
  }

  private static String formatExamples(List<ShippedExampleEntry> shippedExamples) {
    int idWidth =
        shippedExamples.stream().mapToInt(example -> example.id().length()).max().orElse(0);
    int pathWidth =
        shippedExamples.stream()
            .mapToInt(example -> example.suggestedRequestPath().length())
            .max()
            .orElse(0);
    int workspaceModeWidth =
        shippedExamples.stream()
            .mapToInt(example -> example.workspaceMode().name().length())
            .max()
            .orElse(0);
    int labelIdWidth = Math.max(idWidth, "Example Id".length());
    int labelPathWidth = Math.max(pathWidth, "Suggested Request Path".length());
    int labelWorkspaceModeWidth = Math.max(workspaceModeWidth, "Workspace Mode".length());
    String header =
        ("    %-"
                + labelIdWidth
                + "s  %-"
                + labelPathWidth
                + "s  %-"
                + labelWorkspaceModeWidth
                + "s  %s")
            .formatted("Example Id", "Suggested Request Path", "Workspace Mode", "Summary");
    String separator =
        "    "
            + "-".repeat(labelIdWidth)
            + "  "
            + "-".repeat(labelPathWidth)
            + "  "
            + "-".repeat(labelWorkspaceModeWidth)
            + "  "
            + "-".repeat("Summary".length());
    String body =
        shippedExamples.stream()
            .map(
                example ->
                    formatExampleLine(
                        example, labelIdWidth, labelPathWidth, labelWorkspaceModeWidth))
            .map(line -> line + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
    return header + "\n" + separator + "\n" + body;
  }

  private static String formatExampleLine(
      ShippedExampleEntry example, int idWidth, int pathWidth, int workspaceModeWidth) {
    String line =
        ("    %-" + idWidth + "s  %-" + pathWidth + "s  %-" + workspaceModeWidth + "s  %s")
            .formatted(
                example.id(),
                example.suggestedRequestPath(),
                example.workspaceMode().name(),
                example.summary());
    if (example.requiredPaths().isEmpty()) {
      return line;
    }
    return line + "\n" + "      requiredPaths: " + String.join(", ", example.requiredPaths());
  }

  private static String renderCoordinateSystems(CliSurface.CliTableSection section) {
    int leftWidth =
        Math.max(
            section.leftHeader().length(),
            section.entries().stream().mapToInt(entry -> entry.pattern().length()).max().orElse(0));
    String separator =
        "  " + "-".repeat(leftWidth) + " " + "-".repeat(section.rightHeader().length());
    return section.label()
        + ":\n"
        + ("  %-" + leftWidth + "s %s\n").formatted(section.leftHeader(), section.rightHeader())
        + separator
        + "\n"
        + section.entries().stream()
            .map(
                entry ->
                    ("  %-" + leftWidth + "s %s").formatted(entry.pattern(), entry.convention()))
            .map(line -> line + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
  }

  private static String renderSection(CliSurface.CliSection section) {
    return section.label() + ":\n" + indentLinesWrapped(section.lines(), 2);
  }

  private static String renderWorkflows(CliSurface.CliWorkflowSection section) {
    return section.label()
        + ":\n"
        + section.entries().stream()
            .map(entry -> "  " + entry.title() + ":\n" + indentLinesWrapped(entry.lines(), 4))
            .collect(java.util.stream.Collectors.joining("\n"));
  }

  static String renderDefinitions(CliSurface.CliDefinitionSection section) {
    if (section.entries().isEmpty()) {
      return section.label() + ":";
    }
    int width =
        section.entries().stream().mapToInt(entry -> entry.label().length() + 1).max().orElse(0);
    return section.label()
        + ":\n"
        + section.entries().stream()
            .map(entry -> wrappedDefinition(entry, width))
            .map(block -> block + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
  }

  private static String renderCommandExample(
      CliSurface.CliCommandExample section, String containerTag) {
    String commands =
        section.commandLines().stream()
            .map(line -> replacePlaceholders(line, containerTag))
            .map(line -> "  " + line)
            .map(line -> line + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
    if (section.description().isEmpty()) {
      return section.label() + ":\n" + commands;
    }
    return section.label()
        + ":\n"
        + commands
        + "\n\n"
        + indentLinesWrapped(List.of(section.description().orElseThrow()), 2);
  }

  private static String renderDiscovery(
      CliSurface.CliDiscoverySection section, String discoveryExamples) {
    return section.label()
        + ":\n"
        + indentLinesWrapped(section.lines(), 2)
        + "\n"
        + "  "
        + section.builtInExamplesLabel()
        + ":\n"
        + discoveryExamples
        + "\n"
        + "  "
        + section.printOneExampleLabel()
        + ":\n"
        + "    "
        + section.printOneExampleCommand()
        + "\n"
        + indentLinesWrapped(section.protocolCatalogNotes(), 2);
  }

  static String renderReferences(CliSurface.CliReferenceSection section, String documentRef) {
    if (section.entries().isEmpty()) {
      return section.label() + ":";
    }
    return section.label()
        + ":\n"
        + section.entries().stream()
            .map(
                entry ->
                    "  "
                        + entry.label()
                        + ": "
                        + documentRef
                        + "/"
                        + entry.relativePath()
                        + "\n"
                        + wrappedText(entry.description(), "    ", "    ", HELP_TEXT_WIDTH))
            .map(block -> block + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
  }

  private static String indentLinesWrapped(List<String> lines, int indentSpaces) {
    String indent = " ".repeat(indentSpaces);
    return lines.stream()
        .map(line -> wrappedText(line, indent, indent, HELP_TEXT_WIDTH))
        .map(block -> block + "\n")
        .collect(java.util.stream.Collectors.joining())
        .stripTrailing();
  }

  private static String wrappedDefinition(CliSurface.DefinitionEntry entry, int width) {
    String prefix = ("  %-" + width + "s  ").formatted(entry.label() + ":");
    return wrappedText(entry.value(), prefix, " ".repeat(prefix.length()), HELP_TEXT_WIDTH);
  }

  private static String wrappedText(
      String text, String firstPrefix, String continuationPrefix, int width) {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(firstPrefix, "firstPrefix must not be null");
    Objects.requireNonNull(continuationPrefix, "continuationPrefix must not be null");
    String normalizedText = text.trim();
    StringBuilder builder = new StringBuilder(firstPrefix);
    String currentPrefix = firstPrefix;
    int lineLength = currentPrefix.length();
    java.util.regex.Matcher matcher = WHITESPACE_PATTERN.matcher(normalizedText);
    int cursor = 0;
    while (cursor < normalizedText.length()) {
      int nextBoundary = normalizedText.length();
      if (matcher.find(cursor)) {
        nextBoundary = matcher.start();
      }
      String word = normalizedText.substring(cursor, nextBoundary);
      int separatorWidth = lineLength == currentPrefix.length() ? 0 : 1;
      if (lineLength + separatorWidth + word.length() > width
          && lineLength > currentPrefix.length()) {
        builder.append('\n').append(continuationPrefix).append(word);
        currentPrefix = continuationPrefix;
        lineLength = continuationPrefix.length() + word.length();
      } else {
        if (separatorWidth == 1) {
          builder.append(' ');
          lineLength++;
        }
        builder.append(word);
        lineLength += word.length();
      }
      cursor = nextBoundary;
      while (cursor < normalizedText.length()
          && Character.isWhitespace(normalizedText.charAt(cursor))) {
        cursor++;
      }
    }
    return builder.toString();
  }

  private static String replacePlaceholders(String value, String containerTag) {
    return value.replace("{{CONTAINER_TAG}}", containerTag);
  }
}
