package dev.erst.gridgrind.excel;

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
    return (prefix, suffix) -> ExcelTempFiles.createManagedTempFile(rootedTempPath, prefix, suffix);
  }
}
