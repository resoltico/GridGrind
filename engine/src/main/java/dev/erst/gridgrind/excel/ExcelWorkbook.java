package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySnapshot;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * High-level workbook wrapper around Apache POI for creation, loading, saving, and sheet access.
 */
public final class ExcelWorkbook implements AutoCloseable {
  private final ExcelWorkbookContext context;
  private final ExcelWorkbookFormulas formulas;
  private final ExcelWorkbookSheets sheets;
  private final ExcelWorkbookCustomXml customXml;
  private final ExcelWorkbookProtection protection;
  private final ExcelWorkbookNames names;
  private final ExcelWorkbookTables tables;
  private final ExcelWorkbookPivots pivots;
  private final ExcelWorkbookPersistence persistence;
  private boolean mutatedSinceOpen;

  ExcelWorkbook(XSSFWorkbook workbook) {
    this(
        workbook,
        ExcelFormulaRuntime.poi(workbook.getCreationHelper().createFormulaEvaluator()),
        Optional.empty(),
        ExcelOoxmlPackageSecuritySnapshot.none(),
        Optional.empty());
  }

  ExcelWorkbook(XSSFWorkbook workbook, ExcelFormulaEnvironment formulaEnvironment)
      throws IOException {
    this(
        workbook,
        ExcelFormulaRuntime.poi(workbook, formulaEnvironment),
        Optional.empty(),
        ExcelOoxmlPackageSecuritySnapshot.none(),
        Optional.empty());
  }

  ExcelWorkbook(XSSFWorkbook workbook, ExcelFormulaRuntime formulaRuntime) {
    this(
        workbook,
        formulaRuntime,
        Optional.empty(),
        ExcelOoxmlPackageSecuritySnapshot.none(),
        Optional.empty());
  }

  ExcelWorkbook(
      XSSFWorkbook workbook,
      ExcelFormulaRuntime formulaRuntime,
      Optional<Path> sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      Optional<String> sourceEncryptionPassword) {
    this.context =
        new ExcelWorkbookContext(
            workbook, formulaRuntime, sourcePath, loadedPackageSecurity, sourceEncryptionPassword);
    this.formulas = new ExcelWorkbookFormulas(this);
    this.sheets = new ExcelWorkbookSheets(this);
    this.customXml = new ExcelWorkbookCustomXml(this);
    this.protection = new ExcelWorkbookProtection(this);
    this.names = new ExcelWorkbookNames(this);
    this.tables = new ExcelWorkbookTables(this);
    this.pivots = new ExcelWorkbookPivots(this);
    this.persistence = new ExcelWorkbookPersistence(this);
  }

  /** Adapts a POI evaluator into the GridGrind-owned formula runtime seam. */
  ExcelWorkbook(XSSFWorkbook workbook, FormulaEvaluator formulaEvaluator) {
    this(workbook, ExcelFormulaRuntime.poi(formulaEvaluator));
  }

  /** Returns the formula-operation surface for evaluation, cache management, and diagnostics. */
  public ExcelWorkbookFormulas formulas() {
    return formulas;
  }

  /** Returns the grouped sheet lifecycle, visibility, and protection surface. */
  public ExcelWorkbookSheets sheets() {
    return sheets;
  }

  /** Returns the workbook custom-XML mapping surface. */
  public ExcelWorkbookCustomXml customXml() {
    return customXml;
  }

  /** Returns the workbook-level protection surface. */
  public ExcelWorkbookProtection protection() {
    return protection;
  }

  /** Returns the workbook defined-name authoring and inspection surface. */
  public ExcelWorkbookNames names() {
    return names;
  }

  /** Returns the workbook-global table authoring surface. */
  public ExcelWorkbookTables tables() {
    return tables;
  }

  /** Returns the workbook-global pivot-table authoring surface. */
  public ExcelWorkbookPivots pivots() {
    return pivots;
  }

  /** Returns the workbook persistence and source-metadata surface. */
  public ExcelWorkbookPersistence persistence() {
    return persistence;
  }

  /** Returns the named sheet, creating it if necessary. */
  public ExcelSheet getOrCreateSheet(String sheetName) {
    return ExcelWorkbookSheetAccessSupport.getOrCreateSheet(this, sheetName);
  }

  /** Returns an existing sheet. */
  public ExcelSheet sheet(String sheetName) {
    return ExcelWorkbookSheetAccessSupport.sheet(this, sheetName);
  }

  /** Resets only the in-process evaluator cache after workbook mutations. */
  void invalidateFormulaRuntime() {
    context.formulaRuntime().clearCachedResults();
  }

  /** Returns the evaluator environment facts used by formula reads and diagnostics. */
  ExcelFormulaRuntimeContext formulaRuntimeContext() {
    return context.formulaRuntime().context();
  }

  /** Returns the workbook-level summary facts including active and selected sheet state. */
  WorkbookCoreResult.WorkbookSummary workbookSummary() {
    return new ExcelSheetStateController().summarizeWorkbook(this);
  }

  /** Returns the workbook-level protection facts currently stored in the workbook. */
  ExcelWorkbookProtectionSnapshot workbookProtection() {
    return new ExcelSheetStateController().workbookProtection(this);
  }

  /** Returns the summary facts for one sheet, including visibility and protection state. */
  WorkbookSheetResult.SheetSummary sheetSummary(String sheetName) {
    return new ExcelSheetStateController().summarizeSheet(this, sheetName);
  }

  /** Returns factual OOXML package-security state for the current in-memory workbook. */
  ExcelOoxmlPackageSecuritySnapshot packageSecurity() {
    return ExcelWorkbookPersistenceSupport.packageSecurity(this);
  }

  @Override
  public void close() throws IOException {
    ExcelWorkbookPersistenceSupport.close(this);
  }

  /** Returns the mutable XSSF workbook delegate used by workbook-scoped controllers. */
  public XSSFWorkbook xssfWorkbook() {
    return context.workbook();
  }

  ExcelWorkbookContext context() {
    return context;
  }

  boolean wasMutatedSinceOpenInternal() {
    return mutatedSinceOpen;
  }

  void markPackageMutated() {
    mutatedSinceOpen = true;
  }

  Cell requiredCell(String sheetName, String address) {
    return ExcelWorkbookSheetAccessSupport.requiredCell(this, sheetName, address);
  }

  int requiredSheetIndex(String sheetName) {
    return ExcelWorkbookSheetAccessSupport.requiredSheetIndex(this, sheetName);
  }

  /** Returns whether the POI defined name belongs to the requested workbook or sheet scope. */
  boolean scopeMatches(Name candidate, ExcelNamedRangeScope scope) {
    return ExcelWorkbookNamedRangeSupport.scopeMatches(this, candidate, scope);
  }
}
