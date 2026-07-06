package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindRequestSurfaceContractText;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Cross-flag validation for execution-style CLI invocations. */
final class CliExecutionArgumentValidation {
  private CliExecutionArgumentValidation() {}

  static void validateTerminalArguments(
      Optional<Path> requestPath, Optional<Path> executionRootPath, Optional<Path> responsePath) {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    if (requestAndResponseSharePath(requestPath, responsePath)) {
      throw new CliArgumentsException(
          "--response", "--request and --response must not point to the same path");
    }
    if (stdinRequestMissingExecutionRoot(requestPath, executionRootPath)) {
      throw new CliArgumentsException(
          "--execution-root",
          GridGrindRequestSurfaceContractText.stdinExecutionRootRequiredMessage());
    }
    if (fileRequestConflictsWithExecutionRoot(requestPath, executionRootPath)) {
      throw new CliArgumentsException(
          "--execution-root",
          "--execution-root cannot be combined with --request because the request file directory already owns request-root resolution");
    }
  }

  private static boolean requestAndResponseSharePath(
      Optional<Path> requestPath, Optional<Path> responsePath) {
    if (requestPath.isEmpty() || responsePath.isEmpty()) {
      return false;
    }
    return requestPath
        .orElseThrow()
        .toAbsolutePath()
        .equals(responsePath.orElseThrow().toAbsolutePath());
  }

  private static boolean stdinRequestMissingExecutionRoot(
      Optional<Path> requestPath, Optional<Path> executionRootPath) {
    return CliPathArguments.isStandardInputPath(requestPath) && executionRootPath.isEmpty();
  }

  private static boolean fileRequestConflictsWithExecutionRoot(
      Optional<Path> requestPath, Optional<Path> executionRootPath) {
    return requestPath.isPresent()
        && !CliPathArguments.isStandardInputPath(requestPath)
        && executionRootPath.isPresent();
  }
}
