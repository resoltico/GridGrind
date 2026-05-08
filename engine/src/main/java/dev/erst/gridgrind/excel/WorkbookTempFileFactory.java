package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.nio.file.Path;

/** Temporary-file factory for engine-owned workbook materialization and persistence flows. */
@FunctionalInterface
public interface WorkbookTempFileFactory {
  /** Creates one temporary file owned by the caller. */
  Path createTempFile(String prefix, String suffix) throws IOException;
}
