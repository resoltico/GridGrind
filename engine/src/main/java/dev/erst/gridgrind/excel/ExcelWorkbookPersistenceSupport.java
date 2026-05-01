package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Sheet;

/** Save, close, and package-security support for {@link ExcelWorkbook}. */
final class ExcelWorkbookPersistenceSupport {
  private ExcelWorkbookPersistenceSupport() {}

  static void save(
      ExcelWorkbook workbook,
      Path workbookPath,
      ExcelOoxmlPersistenceOptions persistenceOptions,
      ExcelOoxmlPackageSecuritySupport.TempFileFactory tempFileFactory)
      throws IOException {
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    if ((persistenceOptions == null || persistenceOptions.isEmpty())
        && !workbook.context().loadedPackageSecurity().isSecure()) {
      savePlainWorkbook(workbook, workbookPath);
      return;
    }
    ExcelOoxmlPackageSecuritySupport.saveWorkbook(
        workbook, workbookPath, persistenceOptions, tempFileFactory);
  }

  static void savePlainWorkbook(ExcelWorkbook workbook, Path workbookPath) throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");

    Path absolutePath = workbookPath.toAbsolutePath();
    Files.createDirectories(absolutePath.getParent());
    for (Sheet sheet : workbook.context().workbook()) {
      ExcelRowColumnStructureController.canonicalizeColumnDefinitions(
          (org.apache.poi.xssf.usermodel.XSSFSheet) sheet);
    }
    ExcelTableHeaderSyncSupport.syncAllHeaders(workbook.context().workbook());

    try (OutputStream outputStream = Files.newOutputStream(absolutePath)) {
      workbook.context().workbook().write(outputStream);
    }
  }

  static ExcelOoxmlPackageSecuritySnapshot packageSecurity(ExcelWorkbook workbook) {
    return workbook.wasMutatedSinceOpen()
        ? workbook.context().loadedPackageSecurity().afterMutation()
        : workbook.context().loadedPackageSecurity();
  }

  static void close(ExcelWorkbook workbook) throws IOException {
    IOException failure = null;
    try {
      workbook.context().formulaRuntime().close();
    } catch (IOException exception) {
      failure = exception;
    }
    try {
      workbook.context().workbook().close();
    } catch (IOException exception) {
      if (failure == null) {
        failure = exception;
      } else {
        failure.addSuppressed(exception);
      }
    }
    if (failure != null) {
      throw failure;
    }
  }
}
