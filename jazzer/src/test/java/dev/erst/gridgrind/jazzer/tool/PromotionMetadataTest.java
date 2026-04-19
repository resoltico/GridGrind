package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Verifies the structural contract for committed promoted Jazzer metadata. */
class PromotionMetadataTest {
  @Test
  void committedMetadataPathsAreProjectRelative() throws IOException {
    Path metadataRoot = JazzerHarness.promotedMetadataRoot(Path.of(""));
    try (Stream<Path> stream = Files.walk(metadataRoot)) {
      for (Path metadataPath :
          stream
              .filter(path -> path.getFileName().toString().endsWith(".json"))
              .sorted()
              .toList()) {
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
    Path projectDirectory = Path.of("").toAbsolutePath().normalize();
    Path metadataRoot = JazzerHarness.promotedMetadataRoot(projectDirectory);
    try (Stream<Path> stream = Files.walk(metadataRoot)) {
      for (Path metadataPath :
          stream
              .filter(path -> path.getFileName().toString().endsWith(".json"))
              .sorted()
              .toList()) {
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
  void promotedMetadataDirectoryContainsOnlyJsonAndTextArtifacts() throws IOException {
    Path metadataRoot = JazzerHarness.promotedMetadataRoot(Path.of(""));
    try (Stream<Path> stream = Files.walk(metadataRoot)) {
      List<Path> unexpectedArtifacts =
          stream
              .filter(Files::isRegularFile)
              .filter(
                  path -> {
                    String fileName = path.getFileName().toString();
                    return !fileName.endsWith(".json") && !fileName.endsWith(".txt");
                  })
              .sorted()
              .toList();

      assertEquals(
          List.of(),
          unexpectedArtifacts,
          "Promoted metadata directories must contain only .json and .txt artifacts.");
    }
  }

  @Test
  void everyReplayTextArtifactIsReferencedByJsonMetadata() throws IOException {
    Path projectDirectory = Path.of("").toAbsolutePath().normalize();
    Path metadataRoot = JazzerHarness.promotedMetadataRoot(projectDirectory);

    Set<Path> referencedReplayTexts = new HashSet<>();
    try (Stream<Path> stream = Files.walk(metadataRoot)) {
      for (Path metadataPath :
          stream
              .filter(path -> path.getFileName().toString().endsWith(".json"))
              .sorted()
              .toList()) {
        PromotionMetadata metadata = JazzerJson.read(metadataPath, PromotionMetadata.class);
        referencedReplayTexts.add(metadata.replayTextPath(projectDirectory));
      }
    }

    List<Path> orphanReplayTexts;
    try (Stream<Path> stream = Files.walk(metadataRoot)) {
      orphanReplayTexts =
          stream
              .filter(path -> path.getFileName().toString().endsWith(".txt"))
              .filter(path -> !referencedReplayTexts.contains(path.toAbsolutePath().normalize()))
              .sorted()
              .toList();
    }

    assertEquals(
        List.of(),
        orphanReplayTexts,
        "Every committed replay text artifact must be referenced by one promoted-metadata entry.");
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
                    throw new UncheckedIOException(exception);
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

  @Test
  void protocolWorkflowPromotedArtifactsUseNeutralCaseNames() throws IOException {
    Path projectDirectory = Path.of("").toAbsolutePath().normalize();
    Path inputDirectory = JazzerHarness.protocolWorkflow().inputDirectory(projectDirectory);
    Path metadataDirectory =
        JazzerHarness.protocolWorkflow().promotedMetadataDirectory(projectDirectory);

    try (Stream<Path> stream = Files.list(inputDirectory)) {
      assertEquals(
          List.of(
              "workflow_case_01.bin",
              "workflow_case_02.bin",
              "workflow_case_03.bin",
              "workflow_case_04.bin",
              "workflow_case_05.bin",
              "workflow_case_06.bin",
              "workflow_case_07.bin",
              "workflow_case_08.bin",
              "workflow_case_09.bin",
              "workflow_case_10.bin",
              "workflow_case_11.bin"),
          stream.map(path -> path.getFileName().toString()).sorted().toList(),
          "Opaque protocol-workflow seeds must use neutral case identifiers.");
    }

    try (Stream<Path> stream = Files.list(metadataDirectory)) {
      assertEquals(
          List.of(
              "workflow_case_01.json",
              "workflow_case_01.txt",
              "workflow_case_02.json",
              "workflow_case_02.txt",
              "workflow_case_03.json",
              "workflow_case_03.txt",
              "workflow_case_04.json",
              "workflow_case_04.txt",
              "workflow_case_05.json",
              "workflow_case_05.txt",
              "workflow_case_06.json",
              "workflow_case_06.txt",
              "workflow_case_07.json",
              "workflow_case_07.txt",
              "workflow_case_08.json",
              "workflow_case_08.txt",
              "workflow_case_09.json",
              "workflow_case_09.txt",
              "workflow_case_10.json",
              "workflow_case_10.txt",
              "workflow_case_11.json",
              "workflow_case_11.txt"),
          stream.map(path -> path.getFileName().toString()).sorted().toList(),
          "Opaque protocol-workflow replay artifacts must use the same neutral case identifiers.");
    }
  }
}
