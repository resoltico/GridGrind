package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Owns workbook-open flows and open-failure cleanup for the Excel workbook boundary. */
final class ExcelWorkbookOpenSupport {
  private ExcelWorkbookOpenSupport() {}

  static ExcelWorkbook openMaterializedWorkbook(
      Path workbookPath,
      Path sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      String sourceEncryptionPassword)
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
      Path sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      String sourceEncryptionPassword)
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
      Path sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      String sourceEncryptionPassword)
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
      Path sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      String sourceEncryptionPassword)
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
