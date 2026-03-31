package dev.erst.gridgrind.excel;

import java.nio.file.Path;
import java.util.Objects;

/**
 * Captures whether a local file hyperlink path resolved cleanly in the current workbook context.
 */
sealed interface FileHyperlinkResolution
    permits FileHyperlinkResolution.MalformedPath,
        FileHyperlinkResolution.ResolvedPath,
        FileHyperlinkResolution.UnresolvedRelativePath {
  /** Returns the normalized workbook-stored file path string. */
  String path();

  /** Relative file path could not be resolved because the workbook has no filesystem location. */
  record UnresolvedRelativePath(String path) implements FileHyperlinkResolution {
    public UnresolvedRelativePath {
      Objects.requireNonNull(path, "path must not be null");
      if (path.isBlank()) {
        throw new IllegalArgumentException("path must not be blank");
      }
    }
  }

  /** File path resolved to one concrete filesystem path on the current machine. */
  record ResolvedPath(String path, Path resolvedPath) implements FileHyperlinkResolution {
    public ResolvedPath {
      Objects.requireNonNull(path, "path must not be null");
      Objects.requireNonNull(resolvedPath, "resolvedPath must not be null");
      if (path.isBlank()) {
        throw new IllegalArgumentException("path must not be blank");
      }
      resolvedPath = resolvedPath.toAbsolutePath().normalize();
    }
  }

  /** File path string is not valid for resolution on the current runtime. */
  record MalformedPath(String path, String reason) implements FileHyperlinkResolution {
    public MalformedPath {
      Objects.requireNonNull(path, "path must not be null");
      Objects.requireNonNull(reason, "reason must not be null");
      if (path.isBlank()) {
        throw new IllegalArgumentException("path must not be blank");
      }
      if (reason.isBlank()) {
        throw new IllegalArgumentException("reason must not be blank");
      }
    }
  }
}
