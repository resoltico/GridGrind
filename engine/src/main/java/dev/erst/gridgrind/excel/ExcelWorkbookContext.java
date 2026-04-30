package dev.erst.gridgrind.excel;

import java.nio.file.Path;
import java.util.Objects;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Immutable workbook-owned context shared across workbook façades and controller seams. */
final class ExcelWorkbookContext {
  private final XSSFWorkbook workbook;
  private final WorkbookStyleRegistry styleRegistry;
  private final ExcelFormulaRuntime formulaRuntime;
  private final Path sourcePath;
  private final ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity;
  private final String sourceEncryptionPassword;

  ExcelWorkbookContext(
      XSSFWorkbook workbook,
      ExcelFormulaRuntime formulaRuntime,
      Path sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      String sourceEncryptionPassword) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
    this.styleRegistry = new WorkbookStyleRegistry(workbook);
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.sourcePath = sourcePath;
    this.loadedPackageSecurity =
        Objects.requireNonNull(loadedPackageSecurity, "loadedPackageSecurity must not be null");
    this.sourceEncryptionPassword = sourceEncryptionPassword;
  }

  XSSFWorkbook workbook() {
    return workbook;
  }

  WorkbookStyleRegistry styleRegistry() {
    return styleRegistry;
  }

  ExcelFormulaRuntime formulaRuntime() {
    return formulaRuntime;
  }

  Path sourcePath() {
    return sourcePath;
  }

  ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity() {
    return loadedPackageSecurity;
  }

  String sourceEncryptionPassword() {
    return sourceEncryptionPassword;
  }
}
