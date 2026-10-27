package dev.erst.gridgrind.cli.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.regex.Pattern;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

/** Exercises handwritten JSON through intake/doctor and checks delivered local links. */
@Tag("documentation")
class DocumentationTest {
  private static final Path ROOT = Path.of(System.getProperty("gridgrind.repository"));
  private static final String FENCE = Character.toString((char) 96).repeat(3);
  private static final Pattern JSON =
      Pattern.compile("(?ms)^" + FENCE + "json\\s*\\n(.*?)^" + FENCE + "\\s*$");
  private static final Pattern LINK = Pattern.compile("\\[[^\\]]*\\]\\(([^)]+)\\)");

  @Test
  void everyHandwrittenJsonExamplePassesIntakeAndStaticDoctor() throws Exception {
    int examples = 0;
    for (Path file : markdown()) {
      String text = Files.readString(file);
      if (text.startsWith("<!-- Generated")
          || "CHANGELOG_ARCHIVE.md".equals(file.getFileName().toString())
          || file.startsWith(ROOT.resolve("docs/redesign"))) continue;
      var matcher = JSON.matcher(text);
      while (matcher.find()) {
        byte[] input = matcher.group(1).getBytes(StandardCharsets.UTF_8);
        var intake = GridGrindJson.analyzeRequest(input);
        assertTrue(
            intake.isBindable(),
            () -> file + ": " + intake.structuralProblems() + intake.bindingFailures());
        var report =
            GridGrindEngine.requestDoctor()
                .diagnose(intake, RequestInput.requestFile(file.toString()));
        assertTrue(report.valid(), () -> file + ": " + report.problems());
        examples++;
      }
    }
    assertTrue(examples >= 2, "The extraction must exercise the complete documented examples");
    var invalid =
        GridGrindJson.analyzeRequest(
            "{\"protocolVersion\":1,\"source\":{\"type\":\"NEW\"},\"persistence\":{\"type\":\"NONE\"},\"bogus\":true}"
                .getBytes(StandardCharsets.UTF_8));
    assertFalse(invalid.isBindable(), "The intake control must reject an unknown member");
  }

  @Test
  void localFileAndHeadingLinksResolveWithoutAfadMetadata() throws Exception {
    List<String> failures = new ArrayList<>();
    for (Path file : markdown()) {
      String text = Files.readString(file);
      assertFalse(text.lines().anyMatch(line -> line.startsWith("afad:")), file.toString());
      // JSON examples and metadata are data, not prose link markup.
      text = text.replaceAll("(?ms)^" + FENCE + ".*?^" + FENCE + "\\s*$", "");
      var matcher = LINK.matcher(text);
      while (matcher.find()) {
        String link = matcher.group(1);
        if (link.startsWith("https://") || link.startsWith("http://") || link.startsWith("mailto:"))
          continue;
        String[] parts = link.split("#", 2);
        String target = java.net.URLDecoder.decode(parts[0], StandardCharsets.UTF_8);
        Path destination = target.isEmpty() ? file : file.getParent().resolve(target).normalize();
        if (!Files.exists(destination)) {
          failures.add(file + " -> " + link);
          continue;
        }
        if (parts.length == 2
            && destination.toString().endsWith(".md")
            && !anchors(Files.readString(destination)).contains(parts[1])) {
          failures.add(file + " -> " + link + " (missing heading)");
        }
      }
    }
    assertEquals(List.of(), failures);
  }

  @SuppressWarnings(
      "PMD.UseConcurrentHashMap") // Heading occurrence counts are local to one document.
  private static Set<String> anchors(String text) {
    Set<String> anchors = new HashSet<>();
    java.util.Map<String, Integer> counts = new HashMap<>();
    for (String line : text.lines().toList()) {
      if (!line.matches("^#{1,6} .*")) continue;
      String title =
          line.replaceFirst("^#{1,6} +", "")
              .strip()
              .toLowerCase(Locale.ROOT)
              .replaceAll("[^\\p{L}\\p{N}_ -]", "")
              .replace(' ', '-');
      int occurrence = counts.merge(title, 1, Integer::sum) - 1;
      anchors.add(occurrence == 0 ? title : title + "-" + occurrence);
    }
    return anchors;
  }

  private static List<Path> markdown() throws java.io.IOException {
    List<Path> files = new ArrayList<>();
    for (String name : List.of("README.md", "CHANGELOG.md", "PATENTS.md", "core/VOCABULARY.md"))
      files.add(ROOT.resolve(name));
    try (var paths = Files.walk(ROOT.resolve("docs"))) {
      paths.filter(path -> path.toString().endsWith(".md")).sorted().forEach(files::add);
    }
    return files;
  }
}
