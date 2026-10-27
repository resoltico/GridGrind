package dev.erst.gridgrind.engine.backend.xssf;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Objects;

/** Temporary-file factory for engine-owned workbook materialization and persistence flows. */
@FunctionalInterface
public interface WorkbookTempFileFactory {
  /** Creates one temporary file owned by the caller. */
  Path createTempFile(String prefix, String suffix) throws IOException;

  /** Returns one temp-file factory rooted at the supplied managed directory. */
  static WorkbookTempFileFactory rooted(Path tempRoot) {
    Path rootedTempPath = Objects.requireNonNull(tempRoot, "tempRoot must not be null");
    return (prefix, suffix) ->
        ManagedTempFiles.createManagedTempFile(rootedTempPath, prefix, suffix);
  }

  /** Releases a reserved temporary file so a CREATE_NEW writer can take ownership. */
  static Path prepareCreateNewTarget(Path reservedPath) throws IOException {
    Objects.requireNonNull(reservedPath, "reservedPath must not be null");
    java.nio.file.Files.deleteIfExists(reservedPath);
    return reservedPath;
  }
}
