package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Path;

/** Shared explicit temp-file ownership for workbook tests that reopen or persist artifacts. */
final class ExcelTempFileFactoryTestSupport {
  private static final WorkbookTempFileFactory SHARED_FACTORY = createSharedFactory();

  private ExcelTempFileFactoryTestSupport() {}

  static WorkbookTempFileFactory tempFileFactory() {
    return SHARED_FACTORY;
  }

  private static WorkbookTempFileFactory createSharedFactory() {
    try {
      Path root = ExcelTempFiles.createManagedTempDirectory("gridgrind-test-factory-root-");
      return WorkbookTempFileFactory.rooted(root);
    } catch (IOException exception) {
      throw new UncheckedIOException("Failed to initialize workbook test temp root", exception);
    }
  }
}
