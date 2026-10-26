package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageFileSupport;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySupport;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
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
      WorkbookArtifactWriteDisposition writeDisposition,
      ExcelOoxmlPersistenceOptions persistenceOptions,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    Objects.requireNonNull(persistenceOptions, "persistenceOptions must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");
    if (persistenceOptions.writesPlaintextUnsigned()
        && !workbook.context().loadedPackageSecurity().isSecure()) {
      savePlainWorkbook(workbook, workbookPath, writeDisposition);
      return;
    }
    ExcelOoxmlPackageSecuritySupport.saveWorkbook(
        workbook, workbookPath, writeDisposition, persistenceOptions, tempFileFactory);
  }

  static void savePlainWorkbook(
      ExcelWorkbook workbook, Path workbookPath, WorkbookArtifactWriteDisposition writeDisposition)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");

    Path absolutePath = workbookPath.toAbsolutePath();
    Files.createDirectories(absolutePath.getParent());
    for (Sheet sheet : workbook.context().workbook()) {
      ExcelRowColumnStructureController.canonicalizeColumnDefinitions(
          (org.apache.poi.xssf.usermodel.XSSFSheet) sheet);
    }
    ExcelTableHeaderSyncSupport.syncAllHeaders(workbook.context().workbook());
    ExcelWorkbookDocumentMetadataSupport.normalizeForSave(workbook.context().workbook());

    try (OutputStream outputStream =
        ExcelOoxmlPackageFileSupport.newWorkbookOutputStream(absolutePath, writeDisposition)) {
      workbook.context().workbook().write(outputStream);
    }
    ExcelDeterministicWorkbookArtifactSupport.normalizeWorkbookPackage(absolutePath);
  }

  static OoxmlPackageSecurityReport packageSecurity(ExcelWorkbook workbook) {
    var loaded = workbook.context().loadedPackageSecurity();
    if (!workbook.persistence().wasMutatedSinceOpen() || loaded.signatures().isEmpty())
      return loaded;
    return new OoxmlPackageSecurityReport(
        loaded.encryption(),
        loaded.signatures().stream()
            .map(
                signature ->
                    new dev.erst.gridgrind.core.model.security.OoxmlSignatureReport(
                        signature.packagePartName(),
                        signature.signer(),
                        signature.state().afterMutation()))
            .toList());
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
