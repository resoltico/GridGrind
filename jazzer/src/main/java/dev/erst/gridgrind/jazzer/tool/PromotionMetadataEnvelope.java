package dev.erst.gridgrind.jazzer.tool;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import java.nio.file.Path;

/**
 * Stable metadata envelope used when refreshing promoted replay artifacts after detail-schema
 * changes.
 *
 * <p>The refresher must not deserialize the volatile replay expectation because that payload is
 * exactly what may have changed. It only needs the stable metadata envelope to locate the promoted
 * input and replay text.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
record PromotionMetadataEnvelope(
    String targetKey,
    String sourcePath,
    String promotedInputPath,
    String promotedAt,
    String replayTextPath) {
  PromotionMetadataEnvelope {
    targetKey = PromotionMetadata.requireNonBlank(targetKey, "targetKey");
    sourcePath = PromotionMetadata.normalizeStoredPath(sourcePath, "sourcePath");
    promotedInputPath =
        PromotionMetadata.normalizeStoredPath(promotedInputPath, "promotedInputPath");
    promotedAt = PromotionMetadata.requireNonBlank(promotedAt, "promotedAt");
    replayTextPath = PromotionMetadata.normalizeStoredPath(replayTextPath, "replayTextPath");
  }

  Path promotedInputPath(Path projectDirectory) {
    return PromotionMetadata.resolveStoredPath(projectDirectory, promotedInputPath);
  }

  Path replayTextPath(Path projectDirectory) {
    return PromotionMetadata.resolveStoredPath(projectDirectory, replayTextPath);
  }
}
