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
  private static final Path DEFAULT_TEMP_ROOT_RELATIVE_PATH = Path.of(".gridgrind", "tmp");

  private CliExecutionBindingsFactory() {}

  static GridGrindRequestInputs create(
      Optional<Path> requestPath,
      Optional<Path> executionRootPath,
      Optional<Path> tempRootPath,
      WorkbookPlan request,
      InputStream stdin)
      throws IOException {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
    Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");
    Path workingDirectory = executionWorkingDirectory(requestPath, executionRootPath);
    Path tempRoot = tempRoot(tempRootPath, workingDirectory);
    if (!GridGrindRequestRequirements.requiresStandardInput(request)) {
      return new GridGrindRequestInputs(workingDirectory, tempRoot);
    }
    return new GridGrindRequestInputs(workingDirectory, tempRoot, stdin.readAllBytes());
  }

  static Path executionWorkingDirectory(
      Optional<Path> requestPath, Optional<Path> executionRootPath) {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
    if (requestPath.isEmpty() || CliPathArguments.isStandardInputPath(requestPath)) {
      return executionRootPath
          .orElseThrow(() -> new IllegalArgumentException("executionRootPath must be present"))
          .toAbsolutePath()
          .normalize();
    }
    Path normalizedRequestPath = requestPath.orElseThrow().toAbsolutePath().normalize();
    Path parent = normalizedRequestPath.getParent();
    return parent == null ? normalizedRequestPath : parent;
  }

  static Path tempRoot(Optional<Path> tempRootPath, Path workingDirectory) {
    Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    return tempRootPath
        .map(path -> path.toAbsolutePath().normalize())
        .orElseGet(
            () ->
                workingDirectory
                    .resolve(DEFAULT_TEMP_ROOT_RELATIVE_PATH)
                    .toAbsolutePath()
                    .normalize());
  }
}
