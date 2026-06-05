package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Package-owned temporary-file factory that avoids crowded shared temp roots. */
final class ExcelTempFiles {
  private static final String GRIDGRIND_TEMP_DIR_NAME = "gridgrind";
  private static final String TEMP_DIR_PROPERTY = "java.io.tmpdir";

  private ExcelTempFiles() {}

  static Path createManagedTempFile(String prefix, String suffix) throws IOException {
    return createManagedTempFile(defaultManagedRoot(), prefix, suffix);
  }

  static Path createManagedTempFile(Path tempRoot, String prefix, String suffix)
      throws IOException {
    Objects.requireNonNull(tempRoot, "tempRoot must not be null");
    Objects.requireNonNull(prefix, "prefix must not be null");
    Objects.requireNonNull(suffix, "suffix must not be null");
    Path normalizedRoot = tempRoot.toAbsolutePath().normalize();
    Files.createDirectories(normalizedRoot);
    return Files.createTempFile(normalizedRoot, prefix, suffix);
  }

  static Path createManagedTempDirectory(String prefix) throws IOException {
    return createManagedTempDirectory(defaultManagedRoot(), prefix);
  }

  static Path createManagedTempDirectory(Path tempRoot, String prefix) throws IOException {
    Objects.requireNonNull(tempRoot, "tempRoot must not be null");
    Objects.requireNonNull(prefix, "prefix must not be null");
    Path normalizedRoot = tempRoot.toAbsolutePath().normalize();
    Files.createDirectories(normalizedRoot);
    return Files.createTempDirectory(normalizedRoot, prefix);
  }

  static @Nullable Path systemTempRoot() {
    String systemTempDir = System.getProperty(TEMP_DIR_PROPERTY);
    return systemTempDir == null || systemTempDir.isBlank()
        ? null
        : Path.of(systemTempDir).resolve(GRIDGRIND_TEMP_DIR_NAME);
  }

  private static Path defaultManagedRoot() throws IOException {
    Path systemTempRoot = systemTempRoot();
    if (systemTempRoot == null) {
      throw new IOException("System temporary-file root is unavailable");
    }
    return systemTempRoot;
  }
}
