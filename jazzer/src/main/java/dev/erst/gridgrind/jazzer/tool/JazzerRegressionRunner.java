package dev.erst.gridgrind.jazzer.tool;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import java.io.BufferedWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

/** Replays one harness's committed promoted inputs directly against GridGrind's replay engine. */
public final class JazzerRegressionRunner {
  private static final String PULSE_PREFIX = "[JAZZER-PULSE] ";

  private JazzerRegressionRunner() {}

  /**
   * Replays the selected harness's committed promoted inputs and exits non-zero on any mismatch.
   */
  public static void main(String[] args) throws IOException {
    System.exit(
        run(
            Path.of("").toAbsolutePath().normalize(),
            parseHarness(args),
            standardWriter(System.out),
            standardWriter(System.err)));
  }

  /** Parses the required `--target <harness-key>` argument pair for direct regression replay. */
  static JazzerHarness parseHarness(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    if (args.length != 2 || !"--target".equals(args[0])) {
      throw new IllegalArgumentException("Usage: JazzerRegressionRunner --target <harness-key>");
    }
    String targetKey = Objects.requireNonNull(args[1], "targetKey must not be null");
    if (targetKey.isBlank()) {
      throw new IllegalArgumentException("targetKey must not be blank");
    }
    return JazzerHarness.fromKey(targetKey);
  }

  /**
   * Replays all committed promoted inputs for one harness and returns a process-style exit code.
   */
  static int run(
      Path projectDirectory,
      JazzerHarness harness,
      PrintWriter outputWriter,
      PrintWriter errorWriter)
      throws IOException {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    Objects.requireNonNull(harness, "harness must not be null");
    Objects.requireNonNull(outputWriter, "outputWriter must not be null");
    Objects.requireNonNull(errorWriter, "errorWriter must not be null");

    List<Path> metadataPaths = promotedMetadataPaths(projectDirectory, harness);
    if (metadataPaths.isEmpty()) {
      errorWriter.println("No promoted metadata entries were found for harness: " + harness.key());
      return 1;
    }

    outputWriter.println(
        PULSE_PREFIX
            + "regression-target phase=plan target="
            + harness.key()
            + " total-inputs="
            + metadataPaths.size());

    for (int index = 0; index < metadataPaths.size(); index++) {
      Path metadataPath = metadataPaths.get(index);
      PromotionMetadata metadata = JazzerJson.read(metadataPath, PromotionMetadata.class);
      if (validateMetadata(projectDirectory, harness, metadataPath, metadata, errorWriter) != 0) {
        return 1;
      }
      Path promotedInputPath = metadata.promotedInputPath(projectDirectory);
      ReplayOutcome currentOutcome =
          JazzerReplaySupport.replay(harness, Files.readAllBytes(promotedInputPath));
      String currentOutcomeKind = JazzerReplaySupport.outcomeKind(currentOutcome);
      ReplayExpectation currentExpectation = JazzerReplaySupport.expectationFor(currentOutcome);
      if (!metadata.replayOutcome().equals(currentOutcomeKind)
          || !metadata.expectation().equals(currentExpectation)) {
        writeReplayMismatch(
            errorWriter,
            harness,
            metadataPath,
            promotedInputPath,
            metadata,
            currentOutcome,
            currentOutcomeKind,
            currentExpectation);
        return 1;
      }
      outputWriter.println(
          PULSE_PREFIX
              + "regression-input target="
              + harness.key()
              + " completed="
              + (index + 1)
              + "/"
              + metadataPaths.size()
              + " name="
              + promotedInputPath.getFileName()
              + " status=SUCCESS");
    }
    outputWriter.println(
        PULSE_PREFIX
            + "regression-target phase=finish target="
            + harness.key()
            + " status=SUCCESS");
    return 0;
  }

  static int validateMetadata(
      Path projectDirectory,
      JazzerHarness harness,
      Path metadataPath,
      PromotionMetadata metadata,
      PrintWriter errorWriter) {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    Objects.requireNonNull(harness, "harness must not be null");
    Objects.requireNonNull(metadataPath, "metadataPath must not be null");
    Objects.requireNonNull(metadata, "metadata must not be null");
    Objects.requireNonNull(errorWriter, "errorWriter must not be null");

    if (!metadata.targetKey().equals(harness.key())) {
      errorWriter.println(
          "Promoted metadata target mismatch for "
              + metadataPath.getFileName()
              + ": expected "
              + harness.key()
              + " but was "
              + metadata.targetKey());
      return 1;
    }

    Path promotedInputPath = metadata.promotedInputPath(projectDirectory);
    if (!Files.exists(promotedInputPath)) {
      errorWriter.println("Committed promoted input does not exist: " + promotedInputPath);
      return 1;
    }

    Path replayTextPath = metadata.replayTextPath(projectDirectory);
    if (!Files.exists(replayTextPath)) {
      errorWriter.println("Committed replay text does not exist: " + replayTextPath);
      return 1;
    }

    return 0;
  }

  static void writeReplayMismatch(
      PrintWriter errorWriter,
      JazzerHarness harness,
      Path metadataPath,
      Path promotedInputPath,
      PromotionMetadata metadata,
      ReplayOutcome currentOutcome,
      String currentOutcomeKind,
      ReplayExpectation currentExpectation) {
    Objects.requireNonNull(errorWriter, "errorWriter must not be null");
    Objects.requireNonNull(harness, "harness must not be null");
    Objects.requireNonNull(metadataPath, "metadataPath must not be null");
    Objects.requireNonNull(promotedInputPath, "promotedInputPath must not be null");
    Objects.requireNonNull(metadata, "metadata must not be null");
    Objects.requireNonNull(currentOutcome, "currentOutcome must not be null");
    Objects.requireNonNull(currentOutcomeKind, "currentOutcomeKind must not be null");
    Objects.requireNonNull(currentExpectation, "currentExpectation must not be null");

    errorWriter.println(
        "Regression mismatch for "
            + harness.key()
            + " input "
            + promotedInputPath.getFileName()
            + " (metadata "
            + metadataPath.getFileName()
            + "): expected "
            + metadata.replayOutcome()
            + " / "
            + expectationText(metadata.expectation())
            + " but got "
            + currentOutcomeKind
            + " / "
            + expectationText(currentExpectation));
    if (currentOutcome instanceof ReplayOutcome.UnexpectedFailure unexpectedFailure) {
      errorWriter.println(unexpectedFailure.stackTrace());
    }
  }

  private static String expectationText(ReplayExpectation expectation) {
    try {
      return JazzerJson.toJson(expectation);
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to serialize replay expectation", exception);
    }
  }

  private static List<Path> promotedMetadataPaths(Path projectDirectory, JazzerHarness harness)
      throws IOException {
    Path metadataRoot = harness.promotedMetadataDirectory(projectDirectory);
    if (!Files.isDirectory(metadataRoot)) {
      return List.of();
    }
    try (var stream = Files.walk(metadataRoot)) {
      return stream
          .filter(path -> path.getFileName().toString().endsWith(".json"))
          .sorted()
          .toList();
    }
  }

  private static PrintWriter standardWriter(OutputStream outputStream) {
    return new PrintWriter(
        new BufferedWriter(new OutputStreamWriter(outputStream, StandardCharsets.UTF_8)), true);
  }
}
