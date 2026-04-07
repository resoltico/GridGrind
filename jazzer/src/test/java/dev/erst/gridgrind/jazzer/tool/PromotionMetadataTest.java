package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Verifies the structural contract for committed promoted Jazzer metadata. */
class PromotionMetadataTest {
  @Test
  void committedMetadataPathsAreProjectRelative() throws IOException {
    Path metadataRoot =
        Path.of("src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata");
    try (Stream<Path> stream = Files.walk(metadataRoot)) {
      for (Path metadataPath :
          stream.filter(path -> path.getFileName().toString().endsWith(".json")).sorted().toList()) {
        PromotionMetadata metadata = JazzerJson.read(metadataPath, PromotionMetadata.class);
        assertFalse(Path.of(metadata.sourcePath()).isAbsolute(), "source path must be relative");
        assertFalse(
            Path.of(metadata.promotedInputPath()).isAbsolute(),
            "promoted input path must be relative");
        assertFalse(
            Path.of(metadata.replayTextPath()).isAbsolute(), "replay text path must be relative");
      }
    }
  }

  @Test
  void committedMetadataPathsResolveWithinProjectDirectory() throws IOException {
    Path metadataRoot =
        Path.of("src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata");
    Path projectDirectory = Path.of("").toAbsolutePath().normalize();
    try (Stream<Path> stream = Files.walk(metadataRoot)) {
      for (Path metadataPath :
          stream.filter(path -> path.getFileName().toString().endsWith(".json")).sorted().toList()) {
        PromotionMetadata metadata = JazzerJson.read(metadataPath, PromotionMetadata.class);
        assertTrue(
            Files.exists(metadata.promotedInputPath(projectDirectory)),
            "promoted input must exist for " + metadataPath.getFileName());
        assertTrue(
            Files.exists(metadata.replayTextPath(projectDirectory)),
            "replay text must exist for " + metadataPath.getFileName());
      }
    }
  }

  @Test
  void everyInputFileHasPromotionMetadata() throws IOException {
    Path projectDirectory = Path.of("").toAbsolutePath().normalize();
    List<JazzerHarness> replayableHarnesses =
        Arrays.stream(JazzerHarness.values())
            .filter(
                harness -> {
                  Path dir = harness.inputDirectory(projectDirectory);
                  return Files.isDirectory(dir);
                })
            .toList();

    List<Path> orphans =
        replayableHarnesses.stream()
            .flatMap(
                harness -> {
                  try {
                    return JazzerReportSupport.orphanedInputs(projectDirectory, harness).stream();
                  } catch (IOException exception) {
                    throw new RuntimeException(exception);
                  }
                })
            .sorted()
            .toList();

    assertEquals(
        List.of(),
        orphans,
        "Every committed input file must have a promoted-metadata entry. "
            + "Run 'jazzer/bin/promote <target> <input-path> <name>' for each listed file "
            + "or move it there via that command.");
  }
}
