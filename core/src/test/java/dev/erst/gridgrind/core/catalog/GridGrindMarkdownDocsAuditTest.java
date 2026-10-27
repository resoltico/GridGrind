package dev.erst.gridgrind.core.catalog;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Keeps the development version table aligned with the pinned dependency catalog. */
class GridGrindMarkdownDocsAuditTest {
  @Test
  void developerFoundationVersionsMatchTheCanonicalVersionCatalog() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    String developerGuide = Files.readString(repositoryRoot.resolve("docs/DEVELOPER.md"));
    String versionCatalog = Files.readString(repositoryRoot.resolve("gradle/libs.versions.toml"));

    assertTrue(
        developerGuide.contains(
            "| Jackson Databind | " + catalogVersion(versionCatalog, "jackson-databind") + " |"),
        "docs/DEVELOPER.md must publish the catalog-owned Jackson Databind version");
    assertTrue(
        developerGuide.contains(
            "| JUnit Jupiter | " + catalogVersion(versionCatalog, "junit-jupiter") + " |"),
        "docs/DEVELOPER.md must publish the catalog-owned JUnit version");
    assertTrue(
        developerGuide.contains("| Log4j Core | " + catalogVersion(versionCatalog, "log4j") + " |"),
        "docs/DEVELOPER.md must publish the catalog-owned Log4j version");
    assertTrue(
        developerGuide.contains("Kotlin `" + catalogVersion(versionCatalog, "kotlin") + "`"),
        "docs/DEVELOPER.md must publish the catalog-owned Kotlin version");
  }

  private static String catalogVersion(String versionCatalog, String key) {
    String prefix = key + " = \"";
    return versionCatalog
        .lines()
        .filter(line -> line.startsWith(prefix))
        .map(line -> line.substring(prefix.length(), line.length() - 1))
        .findFirst()
        .orElseThrow(() -> new AssertionError("No version catalog entry for " + key));
  }
}
