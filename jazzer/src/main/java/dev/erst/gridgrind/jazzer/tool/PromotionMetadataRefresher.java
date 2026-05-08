package dev.erst.gridgrind.jazzer.tool;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import dev.erst.gridgrind.jazzer.support.JazzerRunTarget;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Refreshes committed replay metadata from the current deterministic replay engine. */
final class PromotionMetadataRefresher {
  private PromotionMetadataRefresher() {}

  static int refresh(Path projectDirectory, Optional<String> targetKey) throws IOException {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    Objects.requireNonNull(targetKey, "targetKey must not be null");
    List<Path> metadataPaths;
    try (var stream = Files.walk(JazzerHarness.promotedMetadataRoot(projectDirectory))) {
      metadataPaths =
          stream
              .filter(path -> path.getFileName().toString().endsWith(".json"))
              .filter(path -> matchesTarget(path, targetKey))
              .sorted()
              .toList();
    }

    for (Path metadataPath : metadataPaths) {
      refreshEntry(projectDirectory, metadataPath);
    }
    return metadataPaths.size();
  }

  private static boolean matchesTarget(Path path, Optional<String> targetKey) {
    if (targetKey.isEmpty()) {
      return true;
    }
    Path parent = path.getParent();
    return parent != null
        && parent.getFileName() != null
        && parent.getFileName().toString().equals(targetKey.orElseThrow());
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
