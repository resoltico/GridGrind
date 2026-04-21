package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies promoted-metadata refresh across replay-detail schema changes. */
class PromotionMetadataRefresherTest {
  @TempDir Path projectDirectory;

  @Test
  void refresh_rewritesStaleProtocolRequestMetadataWithoutReadingOldReplayDetails()
      throws IOException {
    byte[] input =
        """
        {
          "protocolVersion": "V1",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "assert-sheet",
              "target": { "type": "BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
              "assertion": { "type": "EXPECT_PRESENT" }
            }
          ]
        }
        """
            .getBytes(StandardCharsets.UTF_8);

    Path promotedInputPath =
        JazzerHarness.protocolRequest().inputDirectory(projectDirectory).resolve("seed.json");
    Files.createDirectories(promotedInputPath.getParent());
    Files.write(promotedInputPath, input);

    Path metadataPath =
        JazzerHarness.protocolRequest()
            .promotedMetadataDirectory(projectDirectory)
            .resolve("seed.json");
    Files.createDirectories(metadataPath.getParent());
    Path replayTextPath = metadataPath.resolveSibling("seed.txt");
    Files.writeString(replayTextPath, "stale replay text" + System.lineSeparator());
    Files.writeString(
        metadataPath,
        """
        {
          "targetKey": "protocol-request",
          "sourcePath": "src/original-seed.json",
          "promotedInputPath": "%s",
          "replayOutcome": "SUCCESS",
          "expectation": {
            "outcomeKind": "SUCCESS",
            "details": {
              "type": "PROTOCOL_REQUEST",
              "inputBytes": %d,
              "decodeOutcome": "PARSED",
              "sourceKind": "NEW",
              "persistenceKind": "NONE",
              "operationCount": 0,
              "operationKinds": {},
              "styleKinds": {},
              "readCount": 0,
              "readKinds": {}
            }
          },
          "promotedAt": "2026-04-17T00:00:00Z",
          "replayTextPath": "%s"
        }
        """
            .formatted(
                PromotionMetadata.relativizePath(projectDirectory, promotedInputPath),
                input.length,
                PromotionMetadata.relativizePath(projectDirectory, replayTextPath)));

    assertEquals(1, PromotionMetadataRefresher.refresh(projectDirectory, "protocol-request"));

    PromotionMetadata metadata = JazzerJson.read(metadataPath, PromotionMetadata.class);
    assertEquals("EXPECTED_INVALID", metadata.replayOutcome());
    assertEquals(
        PromotionMetadata.relativizePath(projectDirectory, promotedInputPath),
        metadata.sourcePath());
    assertEquals("EXPECTED_INVALID", metadata.expectation().outcomeKind());
    ProtocolRequestDetails details =
        assertInstanceOf(ProtocolRequestDetails.class, metadata.expectation().details());
    assertEquals(input.length, details.inputBytes());
    assertEquals("INVALID_REQUEST_SHAPE", details.decodeOutcome());
    assertEquals("NOT_PARSED", details.sourceKind());
    assertEquals("NOT_PARSED", details.persistenceKind());
    assertEquals(0, details.operationCount());
    assertEquals(Map.of(), details.operationKinds());
    assertEquals(Map.of(), details.styleKinds());
    assertEquals(0, details.assertionCount());
    assertEquals(Map.of(), details.assertionKinds());
    assertEquals(0, details.readCount());
    assertEquals(Map.of(), details.readKinds());
    String replayText = Files.readString(replayTextPath);
    assertTrue(replayText.contains("Decode Outcome: INVALID_REQUEST_SHAPE"));
    assertTrue(replayText.contains("Source Kind: NOT_PARSED"));
  }
}
