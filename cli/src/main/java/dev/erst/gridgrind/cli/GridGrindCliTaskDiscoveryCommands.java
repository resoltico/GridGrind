package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Example and task-discovery CLI surfaces. */
final class GridGrindCliTaskDiscoveryCommands {
  private GridGrindCliTaskDiscoveryCommands() {}

  static int example(
      CliCommand.PrintExample command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    var example = GridGrindShippedExamples.find(command.lookupId());
    if (example.isEmpty()) {
      String message = CliCatalogCommandSupport.unknownExampleMessage(command.lookupId());
      return CliCatalogPayloadSupport.writeCliFailure(
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
                      + " requestFileName, workspaceMode, and requiredWorkspacePaths.")),
          prettyJson);
    }
    CliCatalogCommandSupport.emitExamplePortabilityWarning(example.get(), stderr);
    return CliCatalogPayloadSupport.writePayload(
        responseWriter,
        "print-example",
        "built-in example request",
        Optional.of("gridgrind --print-example --lookup " + command.lookupId()),
        command.responsePath(),
        stdout,
        stderr,
        GridGrindJsonOutput.writeRequestBytes(example.get().plan(), prettyJson),
        prettyJson);
  }

  static int exampleCatalog(
      CliCommand.PrintExampleCatalog command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return CliCatalogPayloadSupport.writePayload(
        responseWriter,
        "print-example-catalog",
        "example catalog",
        Optional.of("gridgrind --print-example-catalog"),
        command.responsePath(),
        stdout,
        stderr,
        GridGrindCliJson.writeBytes(GridGrindShippedExamples.catalog(), prettyJson),
        prettyJson);
  }

  static int taskCatalog(
      CliCommand.PrintTaskCatalog command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    if (command.lookupId().isEmpty()) {
      return CliCatalogPayloadSupport.writePayload(
          responseWriter,
          "print-task-catalog",
          "task catalog",
          Optional.of("gridgrind --print-task-catalog"),
          command.responsePath(),
          stdout,
          stderr,
          GridGrindCliJson.writeBytes(GridGrindTaskCatalog.catalog(), prettyJson),
          prettyJson);
    }
    String taskFilter = command.lookupId().orElseThrow();
    var entry = GridGrindTaskCatalog.entryFor(taskFilter);
    if (entry.isEmpty()) {
      String message = CliCatalogCommandSupport.unknownTaskMessage(taskFilter);
      return CliCatalogPayloadSupport.writeCliFailure(
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
                      + " know the work you want but not the stable task id.")),
          prettyJson);
    }
    return CliCatalogPayloadSupport.writeRenderedPayload(
        responseWriter,
        "print-task-catalog",
        "task catalog entry",
        Optional.of("gridgrind --print-task-catalog --lookup " + taskFilter),
        command.responsePath(),
        stdout,
        stderr,
        output ->
            GridGrindJsonOutput.writeCatalogLookupResult(
                output, GridGrindTaskCatalog.catalog().protocolVersion(), entry.get(), prettyJson),
        prettyJson);
  }

  static int taskPlan(
      CliCommand.PrintTaskPlan command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    var task = GridGrindTaskCatalog.entryFor(command.lookupId());
    if (task.isEmpty()) {
      String message = CliCatalogCommandSupport.unknownTaskMessage(command.lookupId());
      return CliCatalogPayloadSupport.writeCliFailure(
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
                      + " DASHBOARD or another catalog id.")),
          prettyJson);
    }
    CliCatalogCommandSupport.emitTaskStarterPortabilityWarning(task.get(), stderr);
    return CliCatalogPayloadSupport.writePayload(
        responseWriter,
        "print-task-plan",
        "task starter request",
        Optional.of("gridgrind --print-task-plan --lookup " + command.lookupId()),
        command.responsePath(),
        stdout,
        stderr,
        GridGrindJsonOutput.writeRequestBytes(
            GridGrindTaskPlanner.requestFor(task.get()), prettyJson),
        prettyJson);
  }

  static int taskKeywordMatch(
      CliCommand.PrintTaskKeywordMatch command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    try {
      return CliCatalogPayloadSupport.writePayload(
          responseWriter,
          "print-task-keyword-match",
          "task keyword match report",
          Optional.of("gridgrind --print-task-keyword-match --query \"" + command.query() + "\""),
          command.responsePath(),
          stdout,
          stderr,
          GridGrindCliJson.writeBytes(
              GridGrindTaskKeywordMatcher.reportFor(command.query()), prettyJson),
          prettyJson);
    } catch (IllegalArgumentException exception) {
      return CliCatalogPayloadSupport.writeCliFailure(
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
                      + " non-stop-word term after normalization.")),
          prettyJson);
    }
  }
}
