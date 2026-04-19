package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Objects;

/** Renders the public CLI help text from contract-owned catalog metadata. */
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

    Catalog catalog = GridGrindProtocolCatalog.catalog();
    CliSurface cliSurface = catalog.cliSurface();
    String requestTemplate = requestTemplateText();
    String discoveryExamples = formatExamples(catalog.shippedExamples());
    return """
        %s

        Usage:
          gridgrind [--request <path>] [--response <path>]
          gridgrind --print-request-template
          gridgrind --print-protocol-catalog [--operation <id>]
          gridgrind --print-example <id>
          gridgrind --help | -h
          gridgrind --version

        Execution:
        %s

        Limits:
        %s

        Request:
        %s

        File Workflow:
        %s

        Coordinate Systems:
          Pattern                Convention / Example
        %s

        Minimal Valid Request:
        %s

        stdin Example:
          gridgrind --print-request-template | gridgrind

        Docker File Example:
          docker run --rm -i \\
            -v "$(pwd)":/workdir \\
            -w /workdir \\
            ghcr.io/resoltico/gridgrind:%s \\
            --request request.json \\
            --response response.json

          In Docker, mount the host directory that contains your request and workbook files, then
          set -w to that mount point so every relative path resolves inside the mounted directory.

        Discovery:
        %s
          Built-in generated examples:
        %s
          Print one built-in example:
            gridgrind --print-example WORKBOOK_HEALTH
          The protocol catalog lists each field, whether it is required, and the nested/plain
          type group accepted by polymorphic fields such as target, action, query, value, style,
          and scope.

        Docs:
          Quick reference: %s/docs/QUICK_REFERENCE.md
          Operations reference: %s/docs/OPERATIONS.md
          Error reference: %s/docs/ERRORS.md

        Flags:
          --request <path>                 Read the JSON request from a file instead of stdin.
          --response <path>                Write the JSON response to a file instead of stdout.
          --print-request-template         Print a minimal valid request JSON document.
          --print-protocol-catalog         Print the machine-readable protocol catalog.
          --operation <id>                 With --print-protocol-catalog, print one entry.
          --print-example <id>             Print one built-in generated example request.
          --help, -h                       Print this help text.
          --version                        Print the GridGrind version and description.
          --license                        Print the GridGrind license and third-party notices.
        """
        .formatted(
            productHeader(version, description),
            indentLines(cliSurface.executionLines(), 2),
            indentLines(cliSurface.limitLines(), 2),
            indentLines(cliSurface.requestLines(), 2),
            indentLines(cliSurface.fileWorkflowLines(), 2),
            formatCoordinateSystems(cliSurface.coordinateSystems()),
            indentBlock(requestTemplate),
            containerTag,
            indentLines(cliSurface.discoveryLines(), 2),
            discoveryExamples,
            documentRef,
            documentRef,
            documentRef);
  }

  /** Returns the shared two-line product header used by help and version surfaces. */
  public static String productHeader(String version, String description) {
    Objects.requireNonNull(version, "version must not be null");
    Objects.requireNonNull(description, "description must not be null");
    return "GridGrind " + version + "\n" + description;
  }

  private static String formatExamples(List<ShippedExampleEntry> shippedExamples) {
    int width = shippedExamples.stream().mapToInt(example -> example.id().length()).max().orElse(0);
    return shippedExamples.stream()
        .map(
            example ->
                ("  %-" + width + "s  %s  %s")
                    .formatted(example.id(), "examples/" + example.fileName(), example.summary()))
        .map(line -> line + "\n")
        .collect(java.util.stream.Collectors.joining())
        .stripTrailing();
  }

  private static String formatCoordinateSystems(List<CliSurface.CoordinateSystemEntry> entries) {
    return entries.stream()
        .map(entry -> "  %-22s %s".formatted(entry.pattern(), entry.convention()))
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
}
