package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.stream.Stream;
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
