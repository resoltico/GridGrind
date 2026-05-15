package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySnapshot;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Immutable workbook-owned context shared across workbook façades and controller seams. */
final class ExcelWorkbookContext {
  private final XSSFWorkbook workbook;
  private final WorkbookStyleRegistry styleRegistry;
  private final ExcelFormulaRuntime formulaRuntime;
  private final Optional<Path> sourcePath;
  private final ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity;
  private final Optional<String> sourceEncryptionPassword;

  ExcelWorkbookContext(
      XSSFWorkbook workbook,
      ExcelFormulaRuntime formulaRuntime,
      Optional<Path> sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
    this.styleRegistry = new WorkbookStyleRegistry(workbook);
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.sourcePath = Objects.requireNonNull(sourcePath, "sourcePath must not be null");
    this.loadedPackageSecurity =
        Objects.requireNonNull(loadedPackageSecurity, "loadedPackageSecurity must not be null");
    this.sourceEncryptionPassword =
        Objects.requireNonNull(
            sourceEncryptionPassword, "sourceEncryptionPassword must not be null");
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

  Optional<Path> sourcePath() {
    return sourcePath;
  }

  ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity() {
    return loadedPackageSecurity;
  }

  Optional<String> sourceEncryptionPassword() {
    return sourceEncryptionPassword;
  }
}
