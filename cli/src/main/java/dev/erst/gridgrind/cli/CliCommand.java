package dev.erst.gridgrind.cli;

import java.nio.file.Path;

/** Parsed CLI command model for one GridGrind process invocation. */
sealed interface CliCommand {
  /** Requests that help text be printed to stdout. */
  record Help() implements CliCommand {}

  /** Requests that the version string be printed to stdout. */
  record Version() implements CliCommand {}

  /** Requests protocol execution using stdin/stdout or explicit request/response file paths. */
  record Execute(Path requestPath, Path responsePath) implements CliCommand {}
}
