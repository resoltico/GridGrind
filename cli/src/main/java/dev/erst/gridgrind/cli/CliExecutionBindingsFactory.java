package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import dev.erst.gridgrind.engine.api.GridGrindRequestRequirements;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Builds execution bindings for one CLI request, rooted at the request file when present. */
final class CliExecutionBindingsFactory {
  private CliExecutionBindingsFactory() {}

  static GridGrindRequestInputs create(
      Optional<Path> requestPath, WorkbookPlan request, InputStream stdin) throws IOException {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");
    Path workingDirectory = executionWorkingDirectory(requestPath);
    if (!GridGrindRequestRequirements.requiresStandardInput(request)) {
      return new GridGrindRequestInputs(workingDirectory);
    }
    return new GridGrindRequestInputs(workingDirectory, stdin.readAllBytes());
  }

  static Path executionWorkingDirectory(Optional<Path> requestPath) {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    if (requestPath.isEmpty()) {
      return Path.of("");
    }
    Path normalizedRequestPath = requestPath.orElseThrow().toAbsolutePath().normalize();
    Path parent = normalizedRequestPath.getParent();
    return parent == null ? normalizedRequestPath : parent;
  }
}
