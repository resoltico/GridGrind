package dev.erst.gridgrind.engine.runtime;

import java.nio.file.Path;

/** Test helper for explicit execution-root bindings with one managed scratch root. */
final class ExecutionInputBindingsFixtureSupport {
  private static final Path MANAGED_TEMP_SEGMENT = Path.of(".gridgrind", "tmp");

  private ExecutionInputBindingsFixtureSupport() {}

  static ExecutionInputBindings bindings(Path workingDirectory) {
    return new ExecutionInputBindings(workingDirectory, managedTempRoot(workingDirectory));
  }

  static ExecutionInputBindings bindings(Path workingDirectory, byte[] standardInputBytes) {
    return new ExecutionInputBindings(
        workingDirectory, managedTempRoot(workingDirectory), standardInputBytes);
  }

  private static Path managedTempRoot(Path workingDirectory) {
    return workingDirectory.toAbsolutePath().normalize().resolve(MANAGED_TEMP_SEGMENT);
  }
}
