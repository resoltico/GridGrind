package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.executor.ExecutionInputBindings;
import dev.erst.gridgrind.executor.SourceBackedPlanResolver;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Objects;

/** Builds execution bindings for one CLI request, rooted at the request file when present. */
final class CliExecutionBindingsFactory {
  private CliExecutionBindingsFactory() {}

  static ExecutionInputBindings create(Path requestPath, WorkbookPlan request, InputStream stdin)
      throws IOException {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");
    Path workingDirectory = executionWorkingDirectory(requestPath);
    if (!SourceBackedPlanResolver.requiresStandardInput(request)) {
      return new ExecutionInputBindings(workingDirectory);
    }
    return new ExecutionInputBindings(workingDirectory, stdin.readAllBytes());
  }

  static Path executionWorkingDirectory(Path requestPath) {
    if (requestPath == null) {
      return Path.of("");
    }
    Path normalizedRequestPath = requestPath.toAbsolutePath().normalize();
    Path parent = normalizedRequestPath.getParent();
    return parent == null ? normalizedRequestPath : parent;
  }
}
