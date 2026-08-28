package dev.erst.gridgrind.contract.doc;

import static org.junit.jupiter.api.Assertions.fail;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/**
 * Validates bidirectional consistency of LIM-NNN traceability markers: every {@code // LIM-NNN}
 * code comment must have a matching entry in LIMITATIONS.md, and every LIMITATIONS.md Code field
 * that references {@code // LIM-NNN} must have at least one matching comment in Java source.
 */
class LimitationsTraceabilityTest {
  private static final Pattern LIM_PATTERN = Pattern.compile("//\\s*(LIM-[A-Z0-9]+)");
  private static final Pattern LIM_HEADING_PATTERN = Pattern.compile("^### (LIM-[A-Z0-9]+)");

  @Test
  void allCodeCommentsHaveMatchingLimitationsEntry() throws IOException {
    Path rootDir = requireRootDir();
    Set<String> definedIds = parseLimitationsIds(rootDir);
    Set<String> orphaned = new LinkedHashSet<>();

    for (String id : codeCommentIds(rootDir)) {
      if (!definedIds.contains(id)) {
        orphaned.add(id);
      }
    }

    if (!orphaned.isEmpty()) {
      fail(
          "Orphaned // LIM-NNN comments in Java source have no matching LIMITATIONS.md entry: "
              + orphaned);
    }
  }

  @Test
  void allLimitationsCodeFieldMarkersHaveMatchingCodeComments() throws IOException {
    Path rootDir = requireRootDir();
    Set<String> codeComments = codeCommentIds(rootDir);
    Set<String> missing = new LinkedHashSet<>();

    for (String required : limitationsCodeFieldIds(rootDir)) {
      if (!codeComments.contains(required)) {
        missing.add(required);
      }
    }

    if (!missing.isEmpty()) {
      fail(
          "LIMITATIONS.md Code fields reference // LIM-NNN markers absent from Java source: "
              + missing);
    }
  }

  private static Path requireRootDir() {
    String root = System.getProperty("gridgrind.root.dir");
    if (root == null || root.isBlank()) {
      throw new IllegalStateException(
          "System property 'gridgrind.root.dir' must be set by the Gradle build");
    }
    return Path.of(root);
  }

  private static Set<String> parseLimitationsIds(Path rootDir) throws IOException {
    Path limitationsFile = rootDir.resolve("docs/LIMITATIONS.md");
    Set<String> ids = new LinkedHashSet<>();
    for (String line : Files.readAllLines(limitationsFile)) {
      Matcher matcher = LIM_HEADING_PATTERN.matcher(line);
      if (matcher.find()) {
        ids.add(matcher.group(1));
      }
    }
    return ids;
  }

  private static Set<String> limitationsCodeFieldIds(Path rootDir) throws IOException {
    Path limitationsFile = rootDir.resolve("docs/LIMITATIONS.md");
    Set<String> ids = new LinkedHashSet<>();
    for (String line : Files.readAllLines(limitationsFile)) {
      if (line.startsWith("| **Code**")) {
        Matcher matcher = LIM_PATTERN.matcher(line);
        while (matcher.find()) {
          ids.add(matcher.group(1));
        }
      }
    }
    return ids;
  }

  private static Set<String> codeCommentIds(Path rootDir) throws IOException {
    Set<String> ids = new LinkedHashSet<>();
    for (Path sourceDir : discoverSourceDirs(rootDir)) {
      try (Stream<Path> files = Files.walk(sourceDir)) {
        files
            .filter(p -> p.toString().endsWith(".java"))
            .forEach(
                file -> {
                  try {
                    for (String line : Files.readAllLines(file)) {
                      Matcher matcher = LIM_PATTERN.matcher(line);
                      while (matcher.find()) {
                        ids.add(matcher.group(1));
                      }
                    }
                  } catch (IOException e) {
                    throw new java.io.UncheckedIOException("Failed to read " + file, e);
                  }
                });
      }
    }
    return ids;
  }

  private static List<Path> discoverSourceDirs(Path rootDir) throws IOException {
    try (Stream<Path> children = Files.list(rootDir)) {
      return children
          .filter(Files::isDirectory)
          .map(module -> module.resolve("src/main/java"))
          .filter(Files::isDirectory)
          .filter(p -> !p.toString().contains("/build/"))
          .toList();
    }
  }
}
