package dev.erst.gridgrind.cli;

import java.nio.file.Path;

/** Parsed CLI command model for one GridGrind process invocation. */
sealed interface CliCommand {
  /** Requests that help text be printed to stdout. */
  record Help() implements CliCommand {}

  /** Requests that the version string be printed to stdout. */
  record Version() implements CliCommand {}

  /** Requests that the license text be printed to stdout. */
  record License() implements CliCommand {}

  /** Requests that a minimal valid request JSON document be printed to stdout. */
  record PrintRequestTemplate() implements CliCommand {}

  /** Requests that one built-in generated example request be printed to stdout. */
  record PrintExample(String exampleId) implements CliCommand {}

  /** Requests that the machine-readable task catalog be printed to stdout. */
  record PrintTaskCatalog(String taskFilter) implements CliCommand {}

  /** Requests that one machine-readable starter task plan be printed to stdout. */
  record PrintTaskPlan(String taskId) implements CliCommand {}

  /** Requests that one machine-readable goal-to-task match report be printed to stdout. */
  record PrintGoalPlan(String goal) implements CliCommand {}

  /** Requests that one authored request be linted and summarized without execution. */
  record DoctorRequest(Path requestPath) implements CliCommand {}

  /**
   * Requests that the machine-readable protocol catalog be printed to stdout.
   *
   * <p>When {@code operationFilter} is non-null, only the uniquely matching entry is printed.
   * Duplicate ids must be qualified as {@code <group>:<id>}. When {@code searchQuery} is non-null,
   * a ranked discovery report is printed. When both are null, the full catalog is printed.
   */
  record PrintProtocolCatalog(String operationFilter, String searchQuery) implements CliCommand {}

  /** Requests protocol execution using stdin/stdout or explicit request/response file paths. */
  record Execute(Path requestPath, Path responsePath) implements CliCommand {}
}
