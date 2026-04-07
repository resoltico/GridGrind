package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Covers direct committed-input regression replay for one Jazzer harness. */
class JazzerRegressionRunnerTest {
  @Nested
  class ParseHarness {
    @Test
    void parseHarness_returnsHarness_whenArgumentsAreValid() {
      assertEquals(
          JazzerHarness.PROTOCOL_REQUEST,
          JazzerRegressionRunner.parseHarness(new String[] {"--target", "protocol-request"}));
    }

    @Test
    void parseHarness_throwsWhenArgumentsAreMissing() {
      assertThrows(IllegalArgumentException.class, () -> JazzerRegressionRunner.parseHarness(new String[0]));
    }

    @Test
    void parseHarness_throwsWhenTargetIsBlank() {
      assertThrows(
          IllegalArgumentException.class, () -> JazzerRegressionRunner.parseHarness(new String[] {"--target", " "}));
    }
  }

  @Nested
  class Run {
    @TempDir Path projectDirectory;

    @Test
    void run_returnsSuccess_whenPromotedInputMatchesRecordedExpectation() throws Exception {
      StringWriter output = new StringWriter();
      StringWriter errors = new StringWriter();
      writePromotedProtocolRequestMetadata(validProtocolRequestMetadata());

      int exitCode =
          JazzerRegressionRunner.run(
              projectDirectory,
              JazzerHarness.PROTOCOL_REQUEST,
              new PrintWriter(output, true),
              new PrintWriter(errors, true));

      assertEquals(0, exitCode);
      assertTrue(
          output.toString().contains(
              "[JAZZER-PULSE] regression-target phase=plan target=protocol-request total-inputs=1"));
      assertTrue(
          output.toString().contains(
              "[JAZZER-PULSE] regression-input target=protocol-request completed=1/1"));
      assertTrue(
          output.toString().contains(
              "[JAZZER-PULSE] regression-target phase=finish target=protocol-request status=SUCCESS"));
      assertTrue(errors.toString().isBlank());
    }

    @Test
    void run_returnsFailure_whenPromotedInputDriftsFromRecordedExpectation() throws Exception {
      StringWriter output = new StringWriter();
      StringWriter errors = new StringWriter();
      PromotionMetadata metadata = validProtocolRequestMetadata();
      writePromotedProtocolRequestMetadata(
          new PromotionMetadata(
              metadata.targetKey(),
              metadata.sourcePath(),
              metadata.promotedInputPath(),
              "EXPECTED_INVALID",
              metadata.expectation(),
              metadata.promotedAt(),
              metadata.replayTextPath()));

      int exitCode =
          JazzerRegressionRunner.run(
              projectDirectory,
              JazzerHarness.PROTOCOL_REQUEST,
              new PrintWriter(output, true),
              new PrintWriter(errors, true));

      assertEquals(1, exitCode);
      assertTrue(errors.toString().contains("Regression mismatch for protocol-request input"));
    }

    @Test
    void run_replaysNonJsonHarnessesWithoutNativeJazzerProvider() throws Exception {
      StringWriter output = new StringWriter();
      StringWriter errors = new StringWriter();
      byte[] input = new byte[] {1, 2, 3};
      writePromotedMetadata(JazzerHarness.ENGINE_COMMAND_SEQUENCE, "command_sequence.bin", input);

      int exitCode =
          JazzerRegressionRunner.run(
              projectDirectory,
              JazzerHarness.ENGINE_COMMAND_SEQUENCE,
              new PrintWriter(output, true),
              new PrintWriter(errors, true));

      assertEquals(0, exitCode);
      assertTrue(
          output.toString().contains(
              "[JAZZER-PULSE] regression-target phase=plan target=engine-command-sequence total-inputs=1"));
      assertTrue(
          output.toString().contains(
              "[JAZZER-PULSE] regression-input target=engine-command-sequence completed=1/1"));
      assertTrue(errors.toString().isBlank());
    }

    private PromotionMetadata validProtocolRequestMetadata() throws Exception {
      Path inputPath = protocolRequestInputPath();
      Files.createDirectories(inputPath.getParent());
      Files.writeString(inputPath, "{\"source\":{\"type\":\"NEW\"},\"operations\":[]}");
      ReplayOutcome outcome =
          JazzerReplaySupport.replay(
              JazzerHarness.PROTOCOL_REQUEST, Files.readAllBytes(inputPath));
      Path replayTextPath = metadataDirectory().resolve("valid_request.txt");
      Files.createDirectories(replayTextPath.getParent());
      Files.writeString(replayTextPath, "Replay Result" + System.lineSeparator());
      return new PromotionMetadata(
          "protocol-request",
          PromotionMetadata.relativizePath(projectDirectory, inputPath),
          PromotionMetadata.relativizePath(projectDirectory, inputPath),
          JazzerReplaySupport.outcomeKind(outcome),
          JazzerReplaySupport.expectationFor(outcome),
          Instant.now().toString(),
          PromotionMetadata.relativizePath(projectDirectory, replayTextPath));
    }

    private void writePromotedMetadata(JazzerHarness harness, String fileName, byte[] input)
        throws Exception {
      Path inputPath = harness.inputDirectory(projectDirectory).resolve(fileName);
      Files.createDirectories(inputPath.getParent());
      Files.write(inputPath, input);
      ReplayOutcome outcome = JazzerReplaySupport.replay(harness, input);
      String baseName = fileName.contains(".") ? fileName.substring(0, fileName.lastIndexOf('.')) : fileName;
      Path replayTextPath = metadataDirectory(harness).resolve(baseName + ".txt");
      Files.createDirectories(replayTextPath.getParent());
      Files.writeString(replayTextPath, "Replay Result" + System.lineSeparator(), StandardCharsets.UTF_8);
      PromotionMetadata metadata =
          new PromotionMetadata(
              harness.key(),
              PromotionMetadata.relativizePath(projectDirectory, inputPath),
              PromotionMetadata.relativizePath(projectDirectory, inputPath),
              JazzerReplaySupport.outcomeKind(outcome),
              JazzerReplaySupport.expectationFor(outcome),
              Instant.now().toString(),
              PromotionMetadata.relativizePath(projectDirectory, replayTextPath));
      JazzerJson.write(metadataDirectory(harness).resolve(baseName + ".json"), metadata);
    }

    private void writePromotedProtocolRequestMetadata(PromotionMetadata metadata) throws Exception {
      Path metadataPath = metadataDirectory().resolve("valid_request.json");
      Files.createDirectories(metadataPath.getParent());
      JazzerJson.write(metadataPath, metadata);
    }

    private Path protocolRequestInputPath() {
      return projectDirectory.resolve(
          "src/fuzz/resources/dev/erst/gridgrind/jazzer/protocol/ProtocolRequestFuzzTestInputs/readRequest/valid_request.json");
    }

    private Path metadataDirectory() {
      return projectDirectory.resolve(
          "src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata/protocol-request");
    }

    private Path metadataDirectory(JazzerHarness harness) {
      return projectDirectory.resolve(
          "src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata/" + harness.key());
    }
  }
}
