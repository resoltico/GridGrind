package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import dev.erst.gridgrind.engine.api.GridGrindRequestRequirements;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;

/** Builds one CLI execution binding plus one owned private scratch directory. */
final class CliExecutionBindingsFactory {
  private static final String SYSTEM_TEMP_PROPERTY = "java.io.tmpdir";
  private static final String MANAGED_TEMP_ROOT_PREFIX = "gridgrind-cli-";

  private CliExecutionBindingsFactory() {}

  static ManagedRequestInputs create(
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
    ManagedTempRoot tempRoot = createManagedTempRoot(tempRootPath);
    if (!GridGrindRequestRequirements.requiresStandardInput(request)) {
      return new ManagedRequestInputs(
          new GridGrindRequestInputs(workingDirectory, tempRoot.root()), tempRoot.createdParent());
    }
    return new ManagedRequestInputs(
        new GridGrindRequestInputs(workingDirectory, tempRoot.root(), stdin.readAllBytes()),
        tempRoot.createdParent());
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

  static Path tempRootParent(Optional<Path> tempRootPath) throws IOException {
    Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
    if (tempRootPath.isPresent()) {
      return tempRootPath.orElseThrow().toAbsolutePath().normalize();
    }
    return defaultTempRootParent();
  }

  static ManagedTempRoot createManagedTempRoot(Optional<Path> tempRootPath) throws IOException {
    Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
    Path tempRootParent = tempRootParent(tempRootPath);
    boolean parentExisted = Files.exists(tempRootParent);
    Files.createDirectories(tempRootParent);
    Path tempRoot =
        Files.createTempDirectory(tempRootParent, MANAGED_TEMP_ROOT_PREFIX)
            .toAbsolutePath()
            .normalize();
    return new ManagedTempRoot(
        tempRoot, parentExisted ? Optional.empty() : Optional.of(tempRootParent));
  }

  static void deleteTreeIfExists(Path root) {
    if (root == null) {
      return;
    }
    try (Stream<Path> paths = Files.walk(root)) {
      paths.sorted(Comparator.reverseOrder()).forEach(CliExecutionBindingsFactory::deletePath);
    } catch (IOException ignored) {
      // Best-effort cleanup for CLI-owned scratch roots only.
    }
  }

  static void deletePath(Path path) {
    try {
      Files.deleteIfExists(path);
    } catch (IOException ignored) {
      // Best-effort cleanup for CLI-owned scratch roots only.
    }
  }

  private static Path defaultTempRootParent() throws IOException {
    String systemTempDir = System.getProperty(SYSTEM_TEMP_PROPERTY);
    if (systemTempDir == null || systemTempDir.isBlank()) {
      throw new IOException("System temporary-file root is unavailable");
    }
    return Path.of(systemTempDir).toAbsolutePath().normalize();
  }

  /** One CLI-owned execution binding plus enough metadata to clean its scratch root on close. */
  static final class ManagedRequestInputs implements AutoCloseable {
    private final GridGrindRequestInputs inputs;
    private final Optional<Path> createdParent;

    ManagedRequestInputs(GridGrindRequestInputs inputs, Optional<Path> createdParent) {
      this.inputs = Objects.requireNonNull(inputs, "inputs must not be null");
      this.createdParent = Objects.requireNonNull(createdParent, "createdParent must not be null");
    }

    GridGrindRequestInputs inputs() {
      return inputs;
    }

    @Override
    public void close() {
      deleteTreeIfExists(inputs.tempRoot());
      createdParent.ifPresent(CliExecutionBindingsFactory::deletePath);
    }
  }

  record ManagedTempRoot(Path root, Optional<Path> createdParent) {
    ManagedTempRoot {
      Objects.requireNonNull(root, "root must not be null");
      Objects.requireNonNull(createdParent, "createdParent must not be null");
    }
  }
}
