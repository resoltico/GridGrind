package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogCliJson;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
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
        "help",
        "help text",
        Optional.of("gridgrind --help"),
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
        "version",
        "version output",
        Optional.of("gridgrind --version"),
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
        "license",
        "license output",
        Optional.of("gridgrind --license"),
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
        "print-request-template",
        "request template",
        Optional.of("gridgrind --print-request-template"),
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
      String message = CliCatalogCommandSupport.unknownExampleMessage(command.lookupId());
      return writeCliFailure(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "print-example",
              "resolve-lookup",
              Optional.of("--lookup"),
              message,
              List.of("gridgrind --print-example-catalog", "gridgrind --help-guidance"),
              Optional.of(
                  "Use --print-example-catalog first when you need the stable example ids,"
                      + " workspaceMode, and requiredPaths.")));
    }
    CliCatalogCommandSupport.emitExamplePortabilityWarning(example.get(), stderr);
    return writePayload(
        responseWriter,
        "print-example",
        "built-in example request",
        Optional.of("gridgrind --print-example --lookup " + command.lookupId()),
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
        "print-example-catalog",
        "example catalog",
        Optional.of("gridgrind --print-example-catalog"),
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
          "print-task-catalog",
          "task catalog",
          Optional.of("gridgrind --print-task-catalog"),
          command.responsePath(),
          stdout,
          stderr,
          GridGrindCliJson.writeTaskCatalogBytes(GridGrindTaskCatalog.catalog()));
    }
    String taskFilter = command.lookupId().orElseThrow();
    var entry = GridGrindTaskCatalog.entryFor(taskFilter);
    if (entry.isEmpty()) {
      String message = CliCatalogCommandSupport.unknownTaskMessage(taskFilter);
      return writeCliFailure(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "print-task-catalog",
              "resolve-lookup",
              Optional.of("--lookup"),
              message,
              List.of(
                  "gridgrind --print-task-catalog",
                  "gridgrind --print-task-keyword-match --query \"monthly sales dashboard\""),
              Optional.of(
                  "Use --print-task-keyword-match --query \"monthly sales dashboard\" when you"
                      + " know the work you want but not the stable task id.")));
    }
    return writePayload(
        responseWriter,
        "print-task-catalog",
        "task catalog entry",
        Optional.of("gridgrind --print-task-catalog --lookup " + taskFilter),
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
      String message = CliCatalogCommandSupport.unknownTaskMessage(command.lookupId());
      return writeCliFailure(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "print-task-plan",
              "resolve-lookup",
              Optional.of("--lookup"),
              message,
              List.of(
                  "gridgrind --print-task-catalog",
                  "gridgrind --print-task-keyword-match --query \"monthly sales dashboard\""),
              Optional.of(
                  "Resolve one valid task id first, then rerun --print-task-plan --lookup"
                      + " DASHBOARD or another catalog id.")));
    }
    CliCatalogCommandSupport.emitTaskStarterPortabilityWarning(task.get(), stderr);
    return writePayload(
        responseWriter,
        "print-task-plan",
        "task starter request",
        Optional.of("gridgrind --print-task-plan --lookup " + command.lookupId()),
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
          "print-task-keyword-match",
          "task keyword match report",
          Optional.of("gridgrind --print-task-keyword-match --query \"" + command.query() + "\""),
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
              "match-query",
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
        "print-protocol-catalog",
        "protocol catalog",
        Optional.of("gridgrind --print-protocol-catalog --full"),
        command.responsePath(),
        stdout,
        stderr,
        GridGrindJson.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog()));
  }

  static int protocolCatalogIndex(
      CliCommand.PrintProtocolCatalogIndex command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return writePayload(
        responseWriter,
        "print-protocol-catalog",
        "protocol catalog index",
        Optional.of("gridgrind --print-protocol-catalog"),
        command.responsePath(),
        stdout,
        stderr,
        output ->
            ProtocolCatalogCliJson.writeProtocolCatalogIndexReport(
                output, CliCatalogCommandSupport.protocolCatalogIndexReport()));
  }

  static int protocolCatalogSearch(
      CliCommand.PrintProtocolCatalogSearch command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return writePayload(
        responseWriter,
        "print-protocol-catalog",
        "protocol catalog search report",
        Optional.of(
            "gridgrind --print-protocol-catalog --search \"" + command.searchQuery() + "\""),
        command.responsePath(),
        stdout,
        stderr,
        output ->
            ProtocolCatalogCliJson.writeProtocolCatalogSearchReport(
                output, CliCatalogCommandSupport.summarizedSearchReport(command.searchQuery())));
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
              "resolve-lookup",
              Optional.of("--lookup"),
              message,
              matches,
              Optional.of(
                  "Rerun the lookup with one qualified id exactly as listed in suggestions.")));
    }
    var lookupValue = GridGrindProtocolCatalog.lookupValueFor(command.lookupId());
    if (lookupValue.isEmpty()) {
      String message = CliCatalogCommandSupport.unknownOperationMessage(command.lookupId());
      return writeCliFailure(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "print-protocol-catalog",
              "resolve-lookup",
              Optional.of("--lookup"),
              message,
              List.of("gridgrind --print-protocol-catalog --search \"sheet layout\""),
              Optional.of(
                  "Use --search when you know the concept but not the exact lookup id or group.")));
    }
    return writePayload(
        responseWriter,
        "print-protocol-catalog",
        "protocol catalog lookup result",
        Optional.of("gridgrind --print-protocol-catalog --lookup " + command.lookupId()),
        command.responsePath(),
        stdout,
        stderr,
        output ->
            GridGrindJson.writeCatalogLookupResult(
                output, GridGrindProtocolCatalog.catalog().protocolVersion(), lookupValue.get()));
  }

  private static int writePayload(
      CliResponseWriter responseWriter,
      String command,
      String payloadName,
      Optional<String> stdoutSuggestion,
      Optional<java.nio.file.Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      byte[] payload)
      throws IOException {
    return responseWriter.writePayload(
        command, payloadName, stdoutSuggestion, responsePath, stdout, stderr, payload, 0);
  }

  private static int writePayload(
      CliResponseWriter responseWriter,
      String command,
      String payloadName,
      Optional<String> stdoutSuggestion,
      Optional<java.nio.file.Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      OutputRenderer renderer)
      throws IOException {
    ByteArrayOutputStream buffer = new ByteArrayOutputStream();
    renderer.write(buffer);
    return responseWriter.writePayload(
        command,
        payloadName,
        stdoutSuggestion,
        responsePath,
        stdout,
        stderr,
        buffer.toByteArray(),
        0);
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
