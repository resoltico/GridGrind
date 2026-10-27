package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import dev.erst.gridgrind.engine.styles.StyleAdapter;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Immutable workbook-owned context shared across workbook façades and controller seams. */
final class WorkbookContext {
  private final XSSFWorkbook workbook;
  private final StyleAdapter styleRegistry;
  private final FormulaRuntime formulaRuntime;
  private final Optional<Path> sourcePath;
  private final OoxmlPackageSecurityReport loadedPackageSecurity;
  private final Optional<String> sourceEncryptionPassword;

  WorkbookContext(
      XSSFWorkbook workbook,
      FormulaRuntime formulaRuntime,
      Optional<Path> sourcePath,
      OoxmlPackageSecurityReport loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
    this.styleRegistry = new StyleAdapter(workbook);
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

  StyleAdapter styleRegistry() {
    return styleRegistry;
  }

  FormulaRuntime formulaRuntime() {
    return formulaRuntime;
  }

  Optional<Path> sourcePath() {
    return sourcePath;
  }

  OoxmlPackageSecurityReport loadedPackageSecurity() {
    return loadedPackageSecurity;
  }

  Optional<String> sourceEncryptionPassword() {
    return sourceEncryptionPassword;
  }
}
