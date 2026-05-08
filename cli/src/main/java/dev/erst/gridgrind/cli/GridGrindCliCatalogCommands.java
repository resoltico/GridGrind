package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
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
        GridGrindCliProductInfo.helpText(GridGrindCliProductInfo.version())
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
    var example = GridGrindShippedExamples.find(command.exampleId());
    if (example.isEmpty()) {
      String message = unknownExampleMessage(command.exampleId());
      return responseWriter.write(
          command.responsePath(),
          stdout,
          stderr,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              message,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
                  dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument.named(
                      "--print-example")),
              new IllegalArgumentException(message)),
          2);
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
    if (command.taskFilter().isEmpty()) {
      return writePayload(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          GridGrindCliJson.writeTaskCatalogBytes(GridGrindTaskCatalog.catalog()));
    }
    String taskFilter = command.taskFilter().orElseThrow();
    var entry = GridGrindTaskCatalog.entryFor(taskFilter);
    if (entry.isEmpty()) {
      String message = unknownTaskMessage(taskFilter);
      return responseWriter.write(
          command.responsePath(),
          stdout,
          stderr,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              message,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
                  dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument.named(
                      "--task")),
              new IllegalArgumentException(message)),
          2);
    }
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        output -> GridGrindCliJson.writeTaskEntry(output, entry.get()));
  }

  static int taskPlan(
      CliCommand.PrintTaskPlan command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    var task = GridGrindTaskCatalog.entryFor(command.taskId());
    if (task.isEmpty()) {
      String message = unknownTaskMessage(command.taskId());
      return responseWriter.write(
          command.responsePath(),
          stdout,
          stderr,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              message,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
                  dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument.named(
                      "--print-task-plan")),
              new IllegalArgumentException(message)),
          2);
    }
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        GridGrindCliJson.writeTaskPlanTemplateBytes(GridGrindTaskPlanner.planFor(task.get())));
  }

  static int taskKeywordMatch(
      CliCommand.PrintTaskKeywordMatch command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        GridGrindCliJson.writeTaskKeywordMatchReportBytes(
            GridGrindTaskKeywordMatcher.reportFor(command.query())));
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
    List<String> matches = GridGrindProtocolCatalog.matchingLookupIds(command.operationFilter());
    if (matches.size() > 1) {
      String message =
          "Ambiguous operation: "
              + command.operationFilter()
              + ". Use one of: "
              + String.join(", ", matches);
      return responseWriter.write(
          command.responsePath(),
          stdout,
          stderr,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              message,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
                  dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument.named(
                      "--operation")),
              new IllegalArgumentException(message)),
          2);
    }
    var lookupValue = GridGrindProtocolCatalog.lookupValueFor(command.operationFilter());
    if (lookupValue.isEmpty()) {
      String message = unknownOperationMessage(command.operationFilter());
      return responseWriter.write(
          command.responsePath(),
          stdout,
          stderr,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              message,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
                  dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument.named(
                      "--operation")),
              new IllegalArgumentException(message)),
          2);
    }
    return writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        output -> GridGrindJson.writeCatalogLookupValue(output, lookupValue.get()));
  }

  private static String unknownExampleMessage(String exampleId) {
    return suggestedExampleId(exampleId)
        .map(
            suggestion ->
                "Unknown example: "
                    + exampleId
                    + ". Example ids use stable upper-case tokens; did you mean "
                    + suggestion
                    + "? Run gridgrind --help and read the Built-in generated examples section"
                    + " to list valid ids.")
        .orElse(
            "Unknown example: "
                + exampleId
                + ". Run gridgrind --help and read the Built-in generated examples section"
                + " to list valid ids.");
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
                + " requires copied examples/ assets beside the request before execution;"
                + " required paths: "
                + requiredPaths
                + ". Inspect --print-example-catalog or --help for portability details.\n")
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
                    + " gridgrind --print-task-keyword-match <query> to discover a close starter plan.")
        .orElse(
            "Unknown task: "
                + taskId
                + ". Run gridgrind --print-task-catalog to list valid ids or"
                + " gridgrind --print-task-keyword-match <query> to discover a close starter plan.");
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
    return "Unknown operation: "
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
                    || example.fileName().equalsIgnoreCase(exampleId)
                    || exampleStem(example.fileName()).equalsIgnoreCase(exampleId)
                    || normalizeLookupToken(exampleStem(example.fileName()))
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

  private static GridGrindResponse.Failure failure(
      GridGrindProblemCode code,
      String message,
      dev.erst.gridgrind.contract.dto.ProblemContext context,
      Throwable cause) {
    return GridGrindResponses.failure(
        GridGrindProtocolVersion.current(),
        GridGrindProblems.problem(code, message, context, cause));
  }

  /** Renders one command-specific payload into the caller-owned output buffer. */
  @FunctionalInterface
  private interface OutputRenderer {
    /** Writes one command payload into the supplied output stream. */
    void write(OutputStream outputStream) throws IOException;
  }
}
