package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Package-owned temporary-file factory that avoids crowded shared temp roots. */
final class ExcelTempFiles {
  private static final String GRIDGRIND_TEMP_DIR_NAME = "gridgrind";
  private static final String USER_HOME_PROPERTY = "user.home";
  private static final String TEMP_DIR_PROPERTY = "java.io.tmpdir";

  private ExcelTempFiles() {}

  static Path createManagedTempFile(String prefix, String suffix) throws IOException {
    Objects.requireNonNull(prefix, "prefix must not be null");
    Objects.requireNonNull(suffix, "suffix must not be null");

    IOException primaryFailure = null;
    for (Path candidateRoot : candidateRoots()) {
      try {
        Files.createDirectories(candidateRoot);
        return Files.createTempFile(candidateRoot, prefix, suffix);
      } catch (IOException exception) {
        if (primaryFailure == null) {
          primaryFailure = exception;
        } else {
          primaryFailure.addSuppressed(exception);
        }
      }
    }
    if (primaryFailure != null) {
      throw primaryFailure;
    }
    throw new IOException("Failed to initialize any temporary-file root");
  }

  static Path createManagedTempDirectory(String prefix) throws IOException {
    Objects.requireNonNull(prefix, "prefix must not be null");

    IOException primaryFailure = null;
    for (Path candidateRoot : candidateRoots()) {
      try {
        Files.createDirectories(candidateRoot);
        return Files.createTempDirectory(candidateRoot, prefix);
      } catch (IOException exception) {
        if (primaryFailure == null) {
          primaryFailure = exception;
        } else {
          primaryFailure.addSuppressed(exception);
        }
      }
    }
    if (primaryFailure != null) {
      throw primaryFailure;
    }
    throw new IOException("Failed to initialize any temporary-directory root");
  }

  private static Iterable<Path> candidateRoots() {
    List<Path> candidateRoots = new ArrayList<>(2);
    Path systemTempRoot = systemTempRoot();
    if (systemTempRoot != null) {
      candidateRoots.add(systemTempRoot);
    }
    Path userHomeFallbackRoot = userHomeFallbackRoot();
    if (userHomeFallbackRoot != null) {
      candidateRoots.add(userHomeFallbackRoot);
    }
    return candidateRoots;
  }

  private static Path systemTempRoot() {
    String systemTempDir = System.getProperty(TEMP_DIR_PROPERTY);
    return systemTempDir == null || systemTempDir.isBlank()
        ? null
        : Path.of(systemTempDir).resolve(GRIDGRIND_TEMP_DIR_NAME);
  }

  private static Path userHomeFallbackRoot() {
    String userHome = System.getProperty(USER_HOME_PROPERTY);
    return userHome == null || userHome.isBlank()
        ? null
        : Path.of(userHome).resolve("." + GRIDGRIND_TEMP_DIR_NAME).resolve("tmp");
  }
}
