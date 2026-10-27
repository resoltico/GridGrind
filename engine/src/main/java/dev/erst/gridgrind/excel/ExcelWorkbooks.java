package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.ResolvedOpenSecurity;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySupport;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Factory and open surface for materializing workbook wrappers. */
public final class ExcelWorkbooks {
  private ExcelWorkbooks() {}

  /** Creates an empty XLSX workbook. */
  public static ExcelWorkbook create() {
    return new ExcelWorkbook(new XSSFWorkbook());
  }

  /** Wraps one already-materialized POI workbook inside the GridGrind workbook boundary. */
  public static ExcelWorkbook wrap(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    return new ExcelWorkbook(workbook);
  }

  /** Creates an empty XLSX workbook with the supplied formula-evaluation environment. */
  public static ExcelWorkbook create(ResolvedFormulaEnvironment formulaEnvironment)
      throws IOException {
    return new ExcelWorkbook(new XSSFWorkbook(), formulaEnvironment);
  }

  /** Opens an existing workbook file from disk with explicit temp-file ownership. */
  public static ExcelWorkbook open(Path workbookPath, WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    return open(workbookPath, new ResolvedOpenSecurity.Unencrypted(), tempFileFactory);
  }

  /**
   * Opens an existing workbook with explicit package-open settings and a custom temp-file factory.
   */
  public static ExcelWorkbook open(
      Path workbookPath, ResolvedOpenSecurity openOptions, WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    return ExcelOoxmlPackageSecuritySupport.openWorkbook(
        workbookPath, openOptions, tempFileFactory);
  }

  /** Opens one materialized plain OOXML package with explicit source-security metadata. */
  public static ExcelWorkbook openMaterializedWorkbook(
      Path workbookPath,
      Optional<Path> sourcePath,
      OoxmlPackageSecurityReport loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword)
      throws IOException {
    return ExcelWorkbookOpenSupport.openMaterializedWorkbook(
        workbookPath, sourcePath, loadedPackageSecurity, sourceEncryptionPassword);
  }

  /** Opens an existing workbook file from disk with formula evaluation and temp ownership. */
  public static ExcelWorkbook open(
      Path workbookPath,
      ResolvedFormulaEnvironment formulaEnvironment,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    return open(
        workbookPath, formulaEnvironment, new ResolvedOpenSecurity.Unencrypted(), tempFileFactory);
  }

  /**
   * Opens an existing workbook with formula and package-open settings plus a custom temp-file
   * factory.
   */
  public static ExcelWorkbook open(
      Path workbookPath,
      ResolvedFormulaEnvironment formulaEnvironment,
      ResolvedOpenSecurity openOptions,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    return ExcelOoxmlPackageSecuritySupport.openWorkbook(
        workbookPath, formulaEnvironment, openOptions, tempFileFactory);
  }

  /** Opens one materialized plain OOXML package with explicit source-security metadata. */
  public static ExcelWorkbook openMaterializedWorkbook(
      Path workbookPath,
      ResolvedFormulaEnvironment formulaEnvironment,
      Optional<Path> sourcePath,
      OoxmlPackageSecurityReport loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword)
      throws IOException {
    return ExcelWorkbookOpenSupport.openMaterializedWorkbook(
        workbookPath,
        formulaEnvironment,
        sourcePath,
        loadedPackageSecurity,
        sourceEncryptionPassword);
  }

  static void closeWorkbookAfterOpenFailure(XSSFWorkbook workbook, Exception exception)
      throws IOException {
    ExcelWorkbookOpenSupport.closeWorkbookAfterOpenFailure(workbook, exception);
  }
}
