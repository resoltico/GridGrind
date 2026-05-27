package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Discovery, help, and non-executing output commands for the CLI surface. */
final class GridGrindCliCatalogCommands {
  private GridGrindCliCatalogCommands() {}

  static int help(
      CliCommand.Help command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        GridGrindCliProductInfo.helpText(command.topic(), GridGrindCliProductInfo.version())
            .getBytes(StandardCharsets.UTF_8));
  }

  static int version(
      CliCommand.Version command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    String version = GridGrindCliProductInfo.version();
    String description = GridGrindCliProductInfo.description();
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        GridGrindCliProductInfo.productHeader(version, description)
            .getBytes(StandardCharsets.UTF_8));
  }

  static int license(
      CliCommand.License command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        GridGrindCliProductInfo.licenseText(GridGrindCli.class).getBytes(StandardCharsets.UTF_8));
  }

  static int requestTemplate(
      CliCommand.PrintRequestTemplate command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        GridGrindJson.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate()));
  }

  static int example(
      CliCommand.PrintExample command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    var example = GridGrindShippedExamples.find(command.lookupId());
    if (example.isEmpty()) {
      String message = unknownExampleMessage(command.lookupId());
      return writeCliFailure(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "print-example",
              Optional.of("--lookup"),
              message,
              List.of("gridgrind --print-example-catalog", "gridgrind --help-guidance"),
              Optional.of(
                  "Use --print-example-catalog first when you need the stable example ids,"
                      + " workspaceMode, and requiredPaths.")));
    }
    emitExamplePortabilityWarning(example.get(), stderr);
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        GridGrindJson.writeRequestBytes(example.get().plan()));
  }

  static int exampleCatalog(
      CliCommand.PrintExampleCatalog command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        GridGrindCliJson.writeShippedExampleCatalogBytes(GridGrindShippedExamples.catalog()));
  }

  static int taskCatalog(
      CliCommand.PrintTaskCatalog command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    if (command.lookupId().isEmpty()) {
      return writePayload(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          GridGrindCliJson.writeTaskCatalogBytes(GridGrindTaskCatalog.catalog()));
    }
    String taskFilter = command.lookupId().orElseThrow();
    var entry = GridGrindTaskCatalog.entryFor(taskFilter);
    if (entry.isEmpty()) {
      String message = unknownTaskMessage(taskFilter);
      return writeCliFailure(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "print-task-catalog",
              Optional.of("--lookup"),
              message,
              List.of("gridgrind --print-task-catalog", "gridgrind --print-task-keyword-match"),
              Optional.of(
                  "Use --print-task-keyword-match --query <text> when you know the work you want"
                      + " but not the stable task id.")));
    }
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        output ->
            GridGrindJson.writeCatalogLookupResult(
                output, GridGrindTaskCatalog.catalog().protocolVersion(), entry.get()));
  }

  static int taskPlan(
      CliCommand.PrintTaskPlan command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    var task = GridGrindTaskCatalog.entryFor(command.lookupId());
    if (task.isEmpty()) {
      String message = unknownTaskMessage(command.lookupId());
      return writeCliFailure(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "print-task-plan",
              Optional.of("--lookup"),
              message,
              List.of("gridgrind --print-task-catalog", "gridgrind --print-task-keyword-match"),
              Optional.of(
                  "Resolve one valid task id first, then rerun --print-task-plan --lookup <id>.")));
    }
    emitTaskStarterPortabilityWarning(task.get(), stderr);
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        GridGrindJson.writeRequestBytes(GridGrindTaskPlanner.requestFor(task.get())));
  }

  static int taskKeywordMatch(
      CliCommand.PrintTaskKeywordMatch command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    try {
      return writePayload(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          GridGrindCliJson.writeTaskKeywordMatchReportBytes(
              GridGrindTaskKeywordMatcher.reportFor(command.query())));
    } catch (IllegalArgumentException exception) {
      return writeCliFailure(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "print-task-keyword-match",
              Optional.of("--query"),
              Objects.requireNonNullElse(exception.getMessage(), "Invalid keyword query"),
              List.of("gridgrind --print-task-catalog", "gridgrind --help-guidance"),
              Optional.of(
                  "Use a natural-language query that leaves at least one searchable"
                      + " non-stop-word term after normalization.")));
    }
  }

  static int protocolCatalogAll(
      CliCommand.PrintProtocolCatalogAll command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        GridGrindJson.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog()));
  }

  static int protocolCatalogSearch(
      CliCommand.PrintProtocolCatalogSearch command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        output ->
            GridGrindJson.writeCatalogLookupValue(
                output, GridGrindProtocolCatalog.searchCatalog(command.searchQuery())));
  }

  static int protocolCatalogLookup(
      CliCommand.PrintProtocolCatalogLookup command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    List<String> matches = GridGrindProtocolCatalog.matchingLookupIds(command.lookupId());
    if (matches.size() > 1) {
      String message =
          "Ambiguous lookup id: "
              + command.lookupId()
              + ". Use one of: "
              + String.join(", ", matches);
      return writeCliFailure(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "print-protocol-catalog",
              Optional.of("--lookup"),
              message,
              matches,
              Optional.of(
                  "Rerun the lookup with one qualified id exactly as listed in suggestions.")));
    }
    var lookupValue = GridGrindProtocolCatalog.lookupValueFor(command.lookupId());
    if (lookupValue.isEmpty()) {
      String message = unknownOperationMessage(command.lookupId());
      return writeCliFailure(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "print-protocol-catalog",
              Optional.of("--lookup"),
              message,
              List.of("gridgrind --print-protocol-catalog --search <text>"),
              Optional.of(
                  "Use --search when you know the concept but not the exact lookup id or group.")));
    }
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        output ->
            GridGrindJson.writeCatalogLookupResult(
                output, GridGrindProtocolCatalog.catalog().protocolVersion(), lookupValue.get()));
  }

  private static String unknownExampleMessage(String exampleId) {
    return suggestedExampleId(exampleId)
        .map(
            suggestion ->
                "Unknown example: "
                    + exampleId
                    + ". Example ids use stable upper-case tokens; did you mean "
                    + suggestion
                    + "? Run gridgrind --print-example-catalog to list valid ids.")
        .orElse(
            "Unknown example: "
                + exampleId
                + ". Run gridgrind --print-example-catalog to list valid ids.");
  }

  private static void emitExamplePortabilityWarning(
      GridGrindShippedExamples.ShippedExample example, OutputStream stderr) throws IOException {
    Objects.requireNonNull(example, "example must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    if (GridGrindShippedExamples.workspaceModeFor(example.id()).orElseThrow()
        != dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS) {
      return;
    }
    var requirements = GridGrindShippedExamples.requirementsFor(example);
    String requiredPaths = String.join(", ", requirements.requiredPaths());
    stderr.write(
        ("Printed example "
                + example.id()
                + " requires copied asset paths beside the request file before execution;"
                + " required paths: "
                + requiredPaths
                + ". Inspect --print-example-catalog or --help-guidance for portability details.\n")
            .getBytes(StandardCharsets.UTF_8));
    stderr.flush();
  }

  private static void emitTaskStarterPortabilityWarning(
      dev.erst.gridgrind.cli.discovery.TaskEntry task, OutputStream stderr) throws IOException {
    Objects.requireNonNull(task, "task must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    if (task.starter().workspaceMode()
        != dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS) {
      return;
    }
    String requiredPaths = String.join(", ", task.starter().requiredPaths());
    stderr.write(
        ("Printed task starter "
                + task.id()
                + " requires copied asset paths beside the request file before execution;"
                + " required paths: "
                + requiredPaths
                + ". Inspect --print-task-catalog or --help-guidance for starter portability"
                + " details.\n")
            .getBytes(StandardCharsets.UTF_8));
    stderr.flush();
  }

  private static String unknownTaskMessage(String taskId) {
    return suggestedTaskId(taskId)
        .map(
            suggestion ->
                "Unknown task: "
                    + taskId
                    + ". Task ids use stable upper-case tokens; did you mean "
                    + suggestion
                    + "? Run gridgrind --print-task-catalog to list valid ids or"
                    + " gridgrind --print-task-keyword-match --query <text> to discover a close"
                    + " task id before printing its task plan.")
        .orElse(
            "Unknown task: "
                + taskId
                + ". Run gridgrind --print-task-catalog to list valid ids or"
                + " gridgrind --print-task-keyword-match --query <text> to discover a close"
                + " task id before printing its task plan.");
  }

  private static Optional<String> suggestedTaskId(String taskId) {
    String normalizedTaskId = normalizeLookupToken(taskId);
    return GridGrindTaskCatalog.catalog().tasks().stream()
        .map(dev.erst.gridgrind.cli.discovery.TaskEntry::id)
        .filter(
            candidate ->
                candidate.equalsIgnoreCase(taskId)
                    || normalizeLookupToken(candidate).equals(normalizedTaskId))
        .findFirst();
  }

  private static String unknownOperationMessage(String operationId) {
    return "Unknown lookup id: "
        + operationId
        + ". Run gridgrind --print-protocol-catalog --search <text> or"
        + " gridgrind --print-protocol-catalog to discover valid lookup ids.";
  }

  private static Optional<String> suggestedExampleId(String exampleId) {
    String normalizedExampleId = normalizeLookupToken(exampleId);
    return GridGrindShippedExamples.examples().stream()
        .filter(
            example ->
                example.id().equalsIgnoreCase(exampleId)
                    || example.requestFileName().equalsIgnoreCase(exampleId)
                    || exampleStem(example.requestFileName()).equalsIgnoreCase(exampleId)
                    || normalizeLookupToken(exampleStem(example.requestFileName()))
                        .equals(normalizedExampleId))
        .map(GridGrindShippedExamples.ShippedExample::id)
        .findFirst();
  }

  private static String normalizeLookupToken(String value) {
    return value.toUpperCase(Locale.ROOT).replace('-', '_').replace(' ', '_');
  }

  private static String exampleStem(String fileName) {
    return fileName.substring(0, fileName.length() - 5);
  }

  private static int writePayload(
      CliResponseWriter responseWriter,
      Optional<java.nio.file.Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      byte[] payload)
      throws IOException {
    return responseWriter.writePayload(responsePath, stdout, stderr, payload, 0);
  }

  private static int writePayload(
      CliResponseWriter responseWriter,
      Optional<java.nio.file.Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      OutputRenderer renderer)
      throws IOException {
    ByteArrayOutputStream buffer = new ByteArrayOutputStream();
    renderer.write(buffer);
    return responseWriter.writePayload(responsePath, stdout, stderr, buffer.toByteArray(), 0);
  }

  private static int writeCliFailure(
      CliResponseWriter responseWriter,
      Optional<java.nio.file.Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      dev.erst.gridgrind.cli.discovery.CliFailureReport failureReport)
      throws IOException {
    return responseWriter.writeCliFailureReport(responsePath, stdout, stderr, failureReport);
  }

  /** Renders one command-specific payload into the caller-owned output buffer. */
  @FunctionalInterface
  private interface OutputRenderer {
    /** Writes one command payload into the supplied output stream. */
    void write(OutputStream outputStream) throws IOException;
  }
}
