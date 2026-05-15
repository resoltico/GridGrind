package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySnapshot;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Owns workbook-open flows and open-failure cleanup for the Excel workbook boundary. */
final class ExcelWorkbookOpenSupport {
  /** Maximum size of a single decompressed ZIP entry in an xlsx package (LIM-026). */
  static final long MAX_ZIP_ENTRY_SIZE = 100L * 1024 * 1024; // 100 MiB

  static {
    ZipSecureFile.setMinInflateRatio(0.01d); // LIM-026: 1:100 max decompression ratio
    ZipSecureFile.setMaxEntrySize(MAX_ZIP_ENTRY_SIZE); // LIM-026
  }

  private ExcelWorkbookOpenSupport() {}

  static ExcelWorkbook openMaterializedWorkbook(
      Path workbookPath,
      Optional<Path> sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(loadedPackageSecurity, "loadedPackageSecurity must not be null");

    Path absolutePath = workbookPath.toAbsolutePath();
    if (!Files.exists(absolutePath)) {
      throw new WorkbookNotFoundException(absolutePath);
    }

    try (InputStream inputStream = Files.newInputStream(absolutePath)) {
      try {
        return openMaterializedWorkbook(
            new XSSFWorkbook(inputStream),
            sourcePath,
            loadedPackageSecurity,
            sourceEncryptionPassword);
      } catch (NotOfficeXmlFileException exception) {
        throw new IllegalArgumentException("Only .xlsx workbooks are supported", exception);
      }
    }
  }

  static ExcelWorkbook openMaterializedWorkbook(
      Path workbookPath,
      ExcelFormulaEnvironment formulaEnvironment,
      Optional<Path> sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(loadedPackageSecurity, "loadedPackageSecurity must not be null");

    Path absolutePath = workbookPath.toAbsolutePath();
    if (!Files.exists(absolutePath)) {
      throw new WorkbookNotFoundException(absolutePath);
    }

    try (InputStream inputStream = Files.newInputStream(absolutePath)) {
      try {
        return openMaterializedWorkbook(
            new XSSFWorkbook(inputStream),
            formulaEnvironment,
            sourcePath,
            loadedPackageSecurity,
            sourceEncryptionPassword);
      } catch (NotOfficeXmlFileException exception) {
        throw new IllegalArgumentException("Only .xlsx workbooks are supported", exception);
      }
    }
  }

  static ExcelWorkbook openMaterializedWorkbook(
      XSSFWorkbook xssfWorkbook,
      Optional<Path> sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword)
      throws IOException {
    try {
      return new ExcelWorkbook(
          xssfWorkbook,
          ExcelFormulaRuntime.poi(xssfWorkbook.getCreationHelper().createFormulaEvaluator()),
          sourcePath,
          loadedPackageSecurity,
          sourceEncryptionPassword);
    } catch (RuntimeException exception) {
      closeWorkbookAfterOpenFailure(xssfWorkbook, exception);
      throw exception;
    }
  }

  static ExcelWorkbook openMaterializedWorkbook(
      XSSFWorkbook xssfWorkbook,
      ExcelFormulaEnvironment formulaEnvironment,
      Optional<Path> sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword)
      throws IOException {
    try {
      return new ExcelWorkbook(
          xssfWorkbook,
          ExcelFormulaRuntime.poi(xssfWorkbook, formulaEnvironment),
          sourcePath,
          loadedPackageSecurity,
          sourceEncryptionPassword);
    } catch (IOException | RuntimeException exception) {
      closeWorkbookAfterOpenFailure(xssfWorkbook, exception);
      throw exception;
    }
  }

  static void closeWorkbookAfterOpenFailure(XSSFWorkbook workbook, Exception exception)
      throws IOException {
    try {
      workbook.close();
    } catch (IOException closeException) {
      exception.addSuppressed(closeException);
    }
  }
}
