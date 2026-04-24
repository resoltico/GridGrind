package dev.erst.gridgrind.cli;

import java.nio.file.Path;

/** Parsed CLI command model for one GridGrind process invocation. */
sealed interface CliCommand {
  /** Requests that help text be emitted as the primary command output. */
  record Help(Path responsePath) implements CliCommand {}

  /** Requests that the version string be emitted as the primary command output. */
  record Version(Path responsePath) implements CliCommand {}

  /** Requests that the license text be emitted as the primary command output. */
  record License(Path responsePath) implements CliCommand {}

  /** Requests that a minimal valid request JSON document be emitted as the primary output. */
  record PrintRequestTemplate(Path responsePath) implements CliCommand {}

  /** Requests that one built-in generated example request be emitted as the primary output. */
  record PrintExample(String exampleId, Path responsePath) implements CliCommand {}

  /** Requests that the machine-readable task catalog be emitted as the primary output. */
  record PrintTaskCatalog(String taskFilter, Path responsePath) implements CliCommand {}

  /** Requests that one machine-readable starter task plan be emitted as the primary output. */
  record PrintTaskPlan(String taskId, Path responsePath) implements CliCommand {}

  /**
   * Requests that one machine-readable goal-to-task match report be emitted as the primary output.
   */
  record PrintGoalPlan(String goal, Path responsePath) implements CliCommand {}

  /** Requests that one authored request be linted and summarized without execution. */
  record DoctorRequest(Path requestPath, Path responsePath) implements CliCommand {}

  /**
   * Requests that the machine-readable protocol catalog be emitted as the primary output.
   *
   * <p>When {@code operationFilter} is non-null, only the uniquely matching entry is printed.
   * Duplicate ids must be qualified as {@code <group>:<id>}. When {@code searchQuery} is non-null,
   * a ranked discovery report is printed. When both are null, the full catalog is printed.
   */
  record PrintProtocolCatalog(String operationFilter, String searchQuery, Path responsePath)
      implements CliCommand {}

  /** Requests protocol execution using stdin/stdout or explicit request/response file paths. */
  record Execute(Path requestPath, Path responsePath) implements CliCommand {}
}
