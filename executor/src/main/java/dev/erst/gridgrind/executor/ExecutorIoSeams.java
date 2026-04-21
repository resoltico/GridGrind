package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySupport;
import dev.erst.gridgrind.excel.ExcelStreamingWorkbookWriter;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.io.IOException;
import java.nio.file.Path;

/** Functional interface for creating executor-owned temporary workbook files. */
@FunctionalInterface
interface TempFileFactory {
  /** Creates one temporary workbook file path for internal execution workflows. */
  Path createTempFile(String prefix, String suffix) throws IOException;
}

/** Closes one materialized event-read workbook path owned by this executor. */
@FunctionalInterface
interface ReadableWorkbookCloser {
  /** Closes one event-read materialized workbook owned by the executor. */
  void close(ExcelOoxmlPackageSecuritySupport.ReadableWorkbook workbook) throws IOException;
}

/** Best-effort filesystem deletion seam used by executor-owned cleanup tests. */
@FunctionalInterface
interface PathDeleteOperation {
  /** Deletes one executor-owned path when present. */
  void deleteIfExists(Path path) throws IOException;
}

/** Functional interface for closing an ExcelWorkbook after request execution. */
@FunctionalInterface
interface WorkbookCloser {
  /** Closes the workbook, releasing any held resources. */
  void close(ExcelWorkbook workbook) throws IOException;
}

/** Applies the streaming-mode calculation side effect at the execution boundary. */
@FunctionalInterface
interface StreamingCalculationApplier {
  /** Applies streaming-mode calculation state to the low-memory workbook writer. */
  void apply(ExcelStreamingWorkbookWriter writer);
}
