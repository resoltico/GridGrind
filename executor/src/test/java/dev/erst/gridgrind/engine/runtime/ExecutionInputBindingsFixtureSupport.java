package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Objects;

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

  static PreparedBindings preparedBindings(Path workingDirectory) {
    ExecutionInputBindings bindings = bindings(workingDirectory);
    RequestPathAccess access = new RequestPathAccess(workingDirectory, bindings.tempFileFactory());
    return new PreparedBindings(bindings.withRequestPathAccess(access), access);
  }

  private static Path managedTempRoot(Path workingDirectory) {
    return workingDirectory.toAbsolutePath().normalize().resolve(MANAGED_TEMP_SEGMENT);
  }

  record PreparedBindings(ExecutionInputBindings bindings, RequestPathAccess access)
      implements AutoCloseable {
    PreparedBindings {
      Objects.requireNonNull(bindings, "bindings must not be null");
      Objects.requireNonNull(access, "access must not be null");
    }

    @Override
    public void close() throws IOException {
      access.close();
    }
  }
}
