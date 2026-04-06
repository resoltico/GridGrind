package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.DynamicTest;
import org.junit.jupiter.api.TestFactory;

/** Verifies that promoted Jazzer inputs still replay to their recorded semantic contract. */
class PromotionMetadataTest {
  @TestFactory
  Stream<DynamicTest> promotedInputsReplayToTheirRecordedExpectation() throws IOException {
    Path metadataRoot =
        Path.of("src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata");
    try (Stream<Path> stream = Files.walk(metadataRoot)) {
      return stream
          .filter(path -> path.getFileName().toString().endsWith(".json"))
          .sorted()
          .map(this::promotedInputReplayTest)
          .toList()
          .stream();
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

  private DynamicTest promotedInputReplayTest(Path metadataPath) {
    return DynamicTest.dynamicTest(
        metadataPath.getFileName().toString(),
        () -> {
          PromotionMetadata metadata = JazzerJson.read(metadataPath, PromotionMetadata.class);
          Path promotedInputPath = Path.of(metadata.promotedInputPath());
          Path replayTextPath = Path.of(metadata.replayTextPath());
          assertTrue(Files.exists(promotedInputPath), "promoted input must exist");
          assertTrue(Files.exists(replayTextPath), "replay text artifact must exist");

          ReplayOutcome currentOutcome =
              JazzerReplaySupport.replay(
                  JazzerHarness.fromKey(metadata.targetKey()), Files.readAllBytes(promotedInputPath));

          assertEquals(metadata.replayOutcome(), JazzerReplaySupport.outcomeKind(currentOutcome));
          assertEquals(metadata.expectation(), JazzerReplaySupport.expectationFor(currentOutcome));
        });
  }
}
