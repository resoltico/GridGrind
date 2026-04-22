package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Build-failing audit over public Markdown file integrity and release frontmatter. */
class GridGrindMarkdownDocsAuditTest {
  private static final Pattern FRONTMATTER_VERSION_PATTERN =
      Pattern.compile("(?m)^version: \"([^\"]+)\"$");
  private static final Pattern MARKDOWN_LINK_PATTERN =
      Pattern.compile("\\[[^\\]]+\\]\\(([^)]+)\\)");

  @Test
  void versionedMarkdownDocsTrackCurrentReleaseVersion() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    String expectedVersion = releaseVersion(repositoryRoot.resolve("gradle.properties"));

    for (Path path : versionedMarkdownDocs(repositoryRoot)) {
      String contents = Files.readString(path);
      Matcher matcher = FRONTMATTER_VERSION_PATTERN.matcher(contents);
      assertTrue(matcher.find(), () -> path + " must carry version frontmatter");
      assertEquals(
          expectedVersion,
          matcher.group(1),
          () -> path + " must track the current release version in frontmatter");
    }
  }

  @Test
  void publicMarkdownDocsHaveNoBrokenLocalLinks() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    List<String> brokenLinks = new ArrayList<>();

    for (Path path : publicMarkdownDocs(repositoryRoot)) {
      String contents = Files.readString(path);
      Matcher matcher = MARKDOWN_LINK_PATTERN.matcher(contents);
      while (matcher.find()) {
        String target = matcher.group(1);
        if (target.startsWith("http://")
            || target.startsWith("https://")
            || target.startsWith("mailto:")
            || target.startsWith("#")) {
          continue;
        }
        String normalizedTarget = target.split("#", 2)[0];
        if (normalizedTarget.isEmpty()) {
          continue;
        }
        if (normalizedTarget.startsWith("/")) {
          brokenLinks.add(path + " -> " + target + " (absolute filesystem paths are not allowed)");
          continue;
        }
        if (!path.getParent().resolve(normalizedTarget).normalize().toFile().exists()) {
          brokenLinks.add(path + " -> " + target);
        }
      }
    }

    assertTrue(
        brokenLinks.isEmpty(),
        () -> "Public Markdown docs contain broken local links: " + brokenLinks);
  }

  private static List<Path> versionedMarkdownDocs(Path repositoryRoot) throws IOException {
    List<Path> paths = new ArrayList<>();
    paths.add(repositoryRoot.resolve("PATENTS.md"));
    paths.add(repositoryRoot.resolve("jazzer/README.md"));
    try (Stream<Path> docs = Files.walk(repositoryRoot.resolve("docs"))) {
      docs.filter(path -> Files.isRegularFile(path) && path.toString().endsWith(".md"))
          .sorted()
          .forEach(paths::add);
    }
    return List.copyOf(paths);
  }

  private static List<Path> publicMarkdownDocs(Path repositoryRoot) throws IOException {
    List<Path> paths = new ArrayList<>();
    paths.add(repositoryRoot.resolve("README.md"));
    paths.add(repositoryRoot.resolve("CHANGELOG.md"));
    paths.add(repositoryRoot.resolve("PATENTS.md"));
    paths.add(repositoryRoot.resolve("jazzer/README.md"));
    try (Stream<Path> docs = Files.walk(repositoryRoot.resolve("docs"))) {
      docs.filter(path -> Files.isRegularFile(path) && path.toString().endsWith(".md"))
          .sorted()
          .forEach(paths::add);
    }
    return List.copyOf(paths);
  }

  private static String releaseVersion(Path gradlePropertiesPath) throws IOException {
    return Files.readAllLines(gradlePropertiesPath).stream()
        .filter(line -> line.startsWith("version="))
        .findFirst()
        .map(line -> line.substring("version=".length()))
        .orElseThrow(() -> new AssertionError("No version= entry found in gradle.properties"));
  }
}
