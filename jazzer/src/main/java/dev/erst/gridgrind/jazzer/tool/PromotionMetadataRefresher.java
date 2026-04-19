package dev.erst.gridgrind.jazzer.tool;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import dev.erst.gridgrind.jazzer.support.JazzerRunTarget;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/** Refreshes committed replay metadata from the current deterministic replay engine. */
final class PromotionMetadataRefresher {
  private PromotionMetadataRefresher() {}

  static int refresh(Path projectDirectory, String targetKey) throws IOException {
    List<Path> metadataPaths;
    try (var stream = Files.walk(JazzerHarness.promotedMetadataRoot(projectDirectory))) {
      metadataPaths =
          stream
              .filter(path -> path.getFileName().toString().endsWith(".json"))
              .filter(
                  path ->
                      targetKey == null
                          || path.getParent().getFileName().toString().equals(targetKey))
              .sorted()
              .toList();
    }

    for (Path metadataPath : metadataPaths) {
      refreshEntry(projectDirectory, metadataPath);
    }
    return metadataPaths.size();
  }

  static void refreshEntry(Path projectDirectory, Path metadataPath) throws IOException {
    PromotionMetadataEnvelope envelope =
        JazzerJson.read(metadataPath, PromotionMetadataEnvelope.class);
    JazzerRunTarget target = JazzerRunTarget.fromKey(envelope.targetKey());
    Path promotedInputPath = envelope.promotedInputPath(projectDirectory);
    Path replayTextPath = envelope.replayTextPath(projectDirectory);
    ReplayOutcome outcome =
        JazzerReplaySupport.replay(target.replayHarness(), Files.readAllBytes(promotedInputPath));
    Files.writeString(
        replayTextPath,
        JazzerTextRenderer.renderReplay(Path.of(envelope.promotedInputPath()), outcome)
            + System.lineSeparator());
    String refreshedSourcePath = refreshedSourcePath(projectDirectory, envelope);
    JazzerJson.write(
        metadataPath,
        new PromotionMetadata(
            envelope.targetKey(),
            refreshedSourcePath,
            envelope.promotedInputPath(),
            JazzerReplaySupport.outcomeKind(outcome),
            JazzerReplaySupport.expectationFor(outcome),
            envelope.promotedAt(),
            envelope.replayTextPath()));
  }

  private static String refreshedSourcePath(
      Path projectDirectory, PromotionMetadataEnvelope envelope) {
    Path originalSourcePath =
        PromotionMetadata.resolveStoredPath(projectDirectory, envelope.sourcePath());
    if (Files.exists(originalSourcePath)) {
      return envelope.sourcePath();
    }
    return envelope.promotedInputPath();
  }
}
