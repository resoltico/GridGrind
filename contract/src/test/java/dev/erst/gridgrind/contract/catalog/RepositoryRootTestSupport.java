package dev.erst.gridgrind.contract.catalog;

import java.nio.file.Files;
import java.nio.file.Path;

/** Shared repository-root lookup for contract-side public-surface tests. */
final class RepositoryRootTestSupport {
  private RepositoryRootTestSupport() {}

  static Path repositoryRoot() {
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
