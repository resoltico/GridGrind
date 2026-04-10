package dev.erst.gridgrind.jazzer.tool;

import java.nio.file.Path;
import java.util.Objects;

/**
 * Describes one committed regression input promoted from a local Jazzer artifact or corpus file.
 */
public record PromotionMetadata(
    String targetKey,
    String sourcePath,
    String promotedInputPath,
    String replayOutcome,
    ReplayExpectation expectation,
    String promotedAt,
    String replayTextPath) {
  public PromotionMetadata {
    targetKey = requireNonBlank(targetKey, "targetKey");
    sourcePath = normalizeStoredPath(sourcePath, "sourcePath");
    promotedInputPath = normalizeStoredPath(promotedInputPath, "promotedInputPath");
    replayOutcome = requireNonBlank(replayOutcome, "replayOutcome");
    Objects.requireNonNull(expectation, "expectation must not be null");
    promotedAt = requireNonBlank(promotedAt, "promotedAt");
    replayTextPath = normalizeStoredPath(replayTextPath, "replayTextPath");
  }

  /**
   * Normalizes one filesystem path to the project-relative text persisted in promotion metadata.
   */
  static String relativizePath(Path projectDirectory, Path path) {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    Objects.requireNonNull(path, "path must not be null");
    Path normalizedProjectDirectory = projectDirectory.toAbsolutePath().normalize();
    Path normalizedPath = path.toAbsolutePath().normalize();
    return normalizeStoredPath(
        normalizedProjectDirectory.relativize(normalizedPath).toString(), "path");
  }

  /** Resolves the recorded promotion source path against the supplied project directory. */
  Path sourcePath(Path projectDirectory) {
    return resolveStoredPath(projectDirectory, sourcePath);
  }

  /** Resolves the committed promoted-input path against the supplied project directory. */
  Path promotedInputPath(Path projectDirectory) {
    return resolveStoredPath(projectDirectory, promotedInputPath);
  }

  /** Resolves the committed replay-text artifact path against the supplied project directory. */
  Path replayTextPath(Path projectDirectory) {
    return resolveStoredPath(projectDirectory, replayTextPath);
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    String normalized = value.strip();
    if (normalized.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return normalized;
  }

  private static String normalizeStoredPath(String value, String fieldName) {
    String normalized = requireNonBlank(value, fieldName);
    Path storedPath = Path.of(normalized).normalize();
    if (storedPath.isAbsolute()) {
      throw new IllegalArgumentException(fieldName + " must be relative to the project directory");
    }
    String storedPathText = storedPath.toString();
    if (storedPathText.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not resolve to the project directory");
    }
    return storedPathText;
  }

  private static Path resolveStoredPath(Path projectDirectory, String storedPath) {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    return projectDirectory.toAbsolutePath().normalize().resolve(storedPath).normalize();
  }
}
