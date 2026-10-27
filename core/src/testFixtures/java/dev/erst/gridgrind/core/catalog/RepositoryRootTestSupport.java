package dev.erst.gridgrind.core.catalog;

import java.nio.file.Files;
import java.nio.file.Path;

/** Shared repository-root lookup for contract-side public-surface tests. */
public final class RepositoryRootTestSupport {
  private RepositoryRootTestSupport() {}

  /** Finds the repository from the current test working directory and its build settings. */
  public static Path repositoryRoot() {
    Path current = Path.of("").toAbsolutePath().normalize();
    while (current != null) {
      if (Files.isRegularFile(current.resolve("settings.gradle.kts"))) {
        return current;
      }
      current = current.getParent();
    }
    throw new IllegalStateException("Failed to locate repository root from " + Path.of(""));
  }
}
