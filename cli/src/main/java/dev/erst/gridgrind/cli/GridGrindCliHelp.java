package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.ShippedExampleEntry;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Objects;

/** Renders the public CLI help text from CLI-owned discovery metadata. */
public final class GridGrindCliHelp {
  private GridGrindCliHelp() {}

  /** Supplies request-template bytes so tests can cover success and failure rendering paths. */
  @FunctionalInterface
  interface RequestTemplateBytesSupplier {
    /** Returns one UTF-8 request-template byte sequence. */
    byte[] get() throws IOException;
  }

  /** Renders the full CLI help text for one packaged product identity and docs reference root. */
  public static String helpText(
      String version, String description, String documentRef, String containerTag) {
    Objects.requireNonNull(version, "version must not be null");
    Objects.requireNonNull(description, "description must not be null");
    Objects.requireNonNull(documentRef, "documentRef must not be null");
    Objects.requireNonNull(containerTag, "containerTag must not be null");

    CliSurface cliSurface = GridGrindProtocolCatalogCliSurface.CLI_SURFACE;
    String requestTemplate = requestTemplateText();
    String discoveryExamples = formatExamples(GridGrindShippedExamples.catalog().examples());
    return String.join(
            "\n\n",
            productHeader(version, description),
            renderCoreCli(cliSurface),
            renderProtocolGrammar(cliSurface, requestTemplate),
            renderOperatorGuidance(cliSurface, containerTag, discoveryExamples),
            renderReferences(cliSurface.docs(), documentRef))
        + "\n";
  }

  /** Returns the shared two-line product header used by help and version surfaces. */
  public static String productHeader(String version, String description) {
    Objects.requireNonNull(version, "version must not be null");
    Objects.requireNonNull(description, "description must not be null");
    return "GridGrind " + version + "\n" + description;
  }

  private static String formatExamples(List<ShippedExampleEntry> shippedExamples) {
    int idWidth =
        shippedExamples.stream().mapToInt(example -> example.id().length()).max().orElse(0);
    int pathWidth =
        shippedExamples.stream()
            .mapToInt(example -> ("examples/" + example.fileName()).length())
            .max()
            .orElse(0);
    int workspaceModeWidth =
        shippedExamples.stream()
            .mapToInt(example -> example.workspaceMode().name().length())
            .max()
            .orElse(0);
    return shippedExamples.stream()
        .map(example -> formatExampleLine(example, idWidth, pathWidth, workspaceModeWidth))
        .map(line -> line + "\n")
        .collect(java.util.stream.Collectors.joining())
        .stripTrailing();
  }

  private static String formatExampleLine(
      ShippedExampleEntry example, int idWidth, int pathWidth, int workspaceModeWidth) {
    String line =
        ("    %-" + idWidth + "s  %-" + pathWidth + "s  %-" + workspaceModeWidth + "s  %s")
            .formatted(
                example.id(),
                "examples/" + example.fileName(),
                example.workspaceMode().name(),
                example.summary());
    if (example.requiredPaths().isEmpty()) {
      return line;
    }
    return line + "\n" + "      requiredPaths: " + String.join(", ", example.requiredPaths());
  }

  private static String renderCoordinateSystems(CliSurface.CliTableSection section) {
    int leftWidth = 22;
    String separator =
        "  "
            + "-".repeat(section.leftHeader().length())
            + " ".repeat(leftWidth - section.leftHeader().length())
            + " "
            + "-".repeat(section.rightHeader().length());
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
    return section.label() + ":\n" + indentLines(section.lines(), 2);
  }

  private static String renderCoreCli(CliSurface cliSurface) {
    return String.join(
        "\n\n", renderSection(cliSurface.usage()), renderDefinitions(cliSurface.flags()));
  }

  private static String renderProtocolGrammar(CliSurface cliSurface, String requestTemplate) {
    return String.join(
        "\n\n",
        "Protocol Grammar:\n"
            + indentLines(
                List.of(
                    "The sections in this block describe the authoritative CLI and request contract."),
                2),
        renderSection(cliSurface.execution()),
        renderDefinitions(cliSurface.limits()),
        renderSection(cliSurface.request()),
        renderDefinitions(cliSurface.fileWorkflow()),
        renderCoordinateSystems(cliSurface.coordinateSystems()),
        renderTemplate(cliSurface.minimalValidRequest(), requestTemplate));
  }

  private static String renderOperatorGuidance(
      CliSurface cliSurface, String containerTag, String discoveryExamples) {
    return String.join(
        "\n\n",
        "Operator Guidance:\n"
            + indentLines(
                List.of(
                    "The sections in this block are usage aids and playbooks; they do not extend"
                        + " the request grammar."),
                2),
        renderWorkflows(cliSurface.workflows()),
        renderCommandExample(cliSurface.stdinExample(), containerTag),
        renderCommandExample(cliSurface.dockerFileExample(), containerTag),
        renderDiscovery(cliSurface.discovery(), discoveryExamples));
  }

  private static String renderWorkflows(CliSurface.CliWorkflowSection section) {
    return section.label()
        + ":\n"
        + section.entries().stream()
            .map(entry -> "  " + entry.title() + ":\n" + indentLines(entry.lines(), 4))
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
            .map(entry -> ("  %-" + width + "s  %s").formatted(entry.label() + ":", entry.value()))
            .map(line -> line + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
  }

  private static String renderTemplate(
      CliSurface.CliTemplateSection section, String requestTemplate) {
    return section.label() + ":\n" + indentBlock(requestTemplate);
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
        + indentLines(List.of(section.description().orElseThrow()), 2);
  }

  private static String renderDiscovery(
      CliSurface.CliDiscoverySection section, String discoveryExamples) {
    return section.label()
        + ":\n"
        + indentLines(section.lines(), 2)
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
        + indentLines(List.of(section.protocolCatalogNote()), 2);
  }

  static String renderReferences(CliSurface.CliReferenceSection section, String documentRef) {
    if (section.entries().isEmpty()) {
      return section.label() + ":";
    }
    return section.label()
        + ":\n"
        + section.entries().stream()
            .map(entry -> "  " + entry.label() + ": " + documentRef + "/" + entry.relativePath())
            .map(line -> line + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
  }

  private static String indentLines(List<String> lines, int indentSpaces) {
    String indent = " ".repeat(indentSpaces);
    return lines.stream()
        .map(line -> indent + line)
        .map(line -> line + "\n")
        .collect(java.util.stream.Collectors.joining())
        .stripTrailing();
  }

  private static String requestTemplateText() {
    return requestTemplateText(
        () -> GridGrindJson.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate()));
  }

  static String requestTemplateText(RequestTemplateBytesSupplier supplier) {
    Objects.requireNonNull(supplier, "supplier must not be null");
    try {
      return new String(supplier.get(), StandardCharsets.UTF_8);
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to render the built-in request template", exception);
    }
  }

  private static String indentBlock(String value) {
    return value.indent(2).stripTrailing();
  }

  private static String replacePlaceholders(String value, String containerTag) {
    return value.replace("{{CONTAINER_TAG}}", containerTag);
  }
}
