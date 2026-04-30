package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
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
  private boolean mutatedSinceOpen;

  private ExcelWorkbook(XSSFWorkbook workbook) {
    this(
        workbook,
        ExcelFormulaRuntime.poi(workbook.getCreationHelper().createFormulaEvaluator()),
        null,
        ExcelOoxmlPackageSecuritySnapshot.none(),
        null);
  }

  private ExcelWorkbook(XSSFWorkbook workbook, ExcelFormulaEnvironment formulaEnvironment)
      throws IOException {
    this(
        workbook,
        ExcelFormulaRuntime.poi(workbook, formulaEnvironment),
        null,
        ExcelOoxmlPackageSecuritySnapshot.none(),
        null);
  }

  ExcelWorkbook(XSSFWorkbook workbook, ExcelFormulaRuntime formulaRuntime) {
    this(workbook, formulaRuntime, null, ExcelOoxmlPackageSecuritySnapshot.none(), null);
  }

  ExcelWorkbook(
      XSSFWorkbook workbook,
      ExcelFormulaRuntime formulaRuntime,
      Path sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      String sourceEncryptionPassword) {
    this.context =
        new ExcelWorkbookContext(
            workbook, formulaRuntime, sourcePath, loadedPackageSecurity, sourceEncryptionPassword);
    this.formulas = new ExcelWorkbookFormulas(this);
  }

  /** Adapts a POI evaluator into the GridGrind-owned formula runtime seam. */
  ExcelWorkbook(XSSFWorkbook workbook, FormulaEvaluator formulaEvaluator) {
    this(workbook, ExcelFormulaRuntime.poi(formulaEvaluator));
  }

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
  public static ExcelWorkbook create(ExcelFormulaEnvironment formulaEnvironment)
      throws IOException {
    return new ExcelWorkbook(new XSSFWorkbook(), formulaEnvironment);
  }

  /** Returns the formula-operation surface for evaluation, cache management, and diagnostics. */
  public ExcelWorkbookFormulas formulas() {
    return formulas;
  }

  /** Opens an existing workbook file from disk. */
  public static ExcelWorkbook open(Path workbookPath) throws IOException {
    return open(workbookPath, new ExcelOoxmlOpenOptions.Unencrypted());
  }

  /** Opens an existing workbook file from disk with optional OOXML package-open settings. */
  public static ExcelWorkbook open(Path workbookPath, ExcelOoxmlOpenOptions openOptions)
      throws IOException {
    return ExcelOoxmlPackageSecuritySupport.openWorkbook(
        workbookPath, openOptions, ExcelTempFiles::createManagedTempFile);
  }

  /**
   * Opens an existing workbook with explicit package-open settings and a custom temp-file factory.
   */
  public static ExcelWorkbook open(
      Path workbookPath,
      ExcelOoxmlOpenOptions openOptions,
      ExcelOoxmlPackageSecuritySupport.TempFileFactory tempFileFactory)
      throws IOException {
    return ExcelOoxmlPackageSecuritySupport.openWorkbook(
        workbookPath, openOptions, tempFileFactory);
  }

  /** Opens one materialized plain OOXML package with explicit source-security metadata. */
  static ExcelWorkbook openMaterializedWorkbook(
      Path workbookPath,
      Path sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      String sourceEncryptionPassword)
      throws IOException {
    return ExcelWorkbookOpenSupport.openMaterializedWorkbook(
        workbookPath, sourcePath, loadedPackageSecurity, sourceEncryptionPassword);
  }

  /** Opens an existing workbook file from disk with the supplied formula environment. */
  public static ExcelWorkbook open(Path workbookPath, ExcelFormulaEnvironment formulaEnvironment)
      throws IOException {
    return open(workbookPath, formulaEnvironment, null);
  }

  /** Opens an existing workbook file from disk with formula and OOXML package-open settings. */
  public static ExcelWorkbook open(
      Path workbookPath,
      ExcelFormulaEnvironment formulaEnvironment,
      ExcelOoxmlOpenOptions openOptions)
      throws IOException {
    return ExcelOoxmlPackageSecuritySupport.openWorkbook(
        workbookPath, formulaEnvironment, openOptions, ExcelTempFiles::createManagedTempFile);
  }

  /**
   * Opens an existing workbook with formula and package-open settings plus a custom temp-file
   * factory.
   */
  public static ExcelWorkbook open(
      Path workbookPath,
      ExcelFormulaEnvironment formulaEnvironment,
      ExcelOoxmlOpenOptions openOptions,
      ExcelOoxmlPackageSecuritySupport.TempFileFactory tempFileFactory)
      throws IOException {
    return ExcelOoxmlPackageSecuritySupport.openWorkbook(
        workbookPath, formulaEnvironment, openOptions, tempFileFactory);
  }

  /** Opens one materialized plain OOXML package with explicit source-security metadata. */
  static ExcelWorkbook openMaterializedWorkbook(
      Path workbookPath,
      ExcelFormulaEnvironment formulaEnvironment,
      Path sourcePath,
      ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity,
      String sourceEncryptionPassword)
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

  /** Returns the named sheet, creating it if necessary. */
  public ExcelSheet getOrCreateSheet(String sheetName) {
    return ExcelWorkbookSheetAccessSupport.getOrCreateSheet(this, sheetName);
  }

  /** Returns an existing sheet. */
  public ExcelSheet sheet(String sheetName) {
    return ExcelWorkbookSheetAccessSupport.sheet(this, sheetName);
  }

  /** Renames an existing sheet to a new destination name. */
  public ExcelWorkbook renameSheet(String sheetName, String newSheetName) {
    return sheetStateController().renameSheet(this, sheetName, newSheetName);
  }

  /** Deletes an existing sheet from the workbook. A workbook must retain at least one sheet. */
  public ExcelWorkbook deleteSheet(String sheetName) {
    return sheetStateController().deleteSheet(this, sheetName);
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  public ExcelWorkbook moveSheet(String sheetName, int targetIndex) {
    return sheetStateController().moveSheet(this, sheetName, targetIndex);
  }

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  public ExcelWorkbook copySheet(
      String sourceSheetName, String newSheetName, ExcelSheetCopyPosition position) {
    return sheetCopyController().copySheet(this, sourceSheetName, newSheetName, position);
  }

  /** Sets the active sheet and ensures it is selected. */
  public ExcelWorkbook setActiveSheet(String sheetName) {
    return sheetStateController().setActiveSheet(this, sheetName);
  }

  /** Sets the selected visible sheet set and normalizes the active tab into that selection. */
  public ExcelWorkbook setSelectedSheets(List<String> sheetNames) {
    return sheetStateController().setSelectedSheets(this, sheetNames);
  }

  /** Sets one sheet visibility while preserving a visible active selected sheet. */
  public ExcelWorkbook setSheetVisibility(String sheetName, ExcelSheetVisibility visibility) {
    return sheetStateController().setSheetVisibility(this, sheetName, visibility);
  }

  /** Returns factual workbook custom-XML mapping metadata. */
  public List<ExcelCustomXmlMappingSnapshot> customXmlMappings() {
    return customXmlController().mappings(context.workbook());
  }

  /** Exports XML for one existing workbook custom-XML mapping. */
  public ExcelCustomXmlExportSnapshot exportCustomXmlMapping(
      ExcelCustomXmlMappingLocator locator, boolean validateSchema, String encoding) {
    return customXmlController()
        .exportMapping(context.workbook(), locator, validateSchema, encoding);
  }

  /** Imports one XML document into one existing workbook custom-XML mapping. */
  public ExcelWorkbook importCustomXmlMapping(ExcelCustomXmlImportDefinition definition) {
    customXmlController().importMapping(context.workbook(), definition);
    return this;
  }

  /** Enables sheet protection with the exact supported lock flags. */
  public ExcelWorkbook setSheetProtection(
      String sheetName, ExcelSheetProtectionSettings protection) {
    return setSheetProtection(sheetName, protection, null);
  }

  /** Enables sheet protection with the exact supported lock flags. */
  public ExcelWorkbook setSheetProtection(
      String sheetName, ExcelSheetProtectionSettings protection, String password) {
    return sheetStateController().setSheetProtection(this, sheetName, protection, password);
  }

  /** Disables sheet protection entirely. */
  public ExcelWorkbook clearSheetProtection(String sheetName) {
    return sheetStateController().clearSheetProtection(this, sheetName);
  }

  /** Enables workbook-level protection and password hashes with authoritative settings. */
  public ExcelWorkbook setWorkbookProtection(ExcelWorkbookProtectionSettings protection) {
    return sheetStateController().setWorkbookProtection(this, protection);
  }

  /** Clears workbook-level protection and password hashes entirely. */
  public ExcelWorkbook clearWorkbookProtection() {
    return sheetStateController().clearWorkbookProtection(this);
  }

  /** Creates or replaces one named range in workbook or sheet scope. */
  public ExcelWorkbook setNamedRange(ExcelNamedRangeDefinition definition) {
    return ExcelWorkbookNamedRangeSupport.setNamedRange(this, definition);
  }

  /** Deletes one named range from workbook or sheet scope. */
  public ExcelWorkbook deleteNamedRange(String name, ExcelNamedRangeScope scope) {
    return ExcelWorkbookNamedRangeSupport.deleteNamedRange(this, name, scope);
  }

  /** Creates or replaces one workbook-global table definition. */
  public ExcelWorkbook setTable(ExcelTableDefinition definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    tableController().setTable(this, definition);
    return this;
  }

  /** Deletes one existing table by workbook-global name and expected sheet name. */
  public ExcelWorkbook deleteTable(String name, String sheetName) {
    tableController().deleteTable(this, name, sheetName);
    return this;
  }

  /** Creates or replaces one workbook-global pivot-table definition. */
  public ExcelWorkbook setPivotTable(ExcelPivotTableDefinition definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    pivotTableController().setPivotTable(this, definition);
    return this;
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet name. */
  public ExcelWorkbook deletePivotTable(String name, String sheetName) {
    pivotTableController().deletePivotTable(this, name, sheetName);
    return this;
  }

  ExcelWorkbook evaluateAllFormulasInternal() {
    return ExcelWorkbookFormulaSupport.evaluateAllFormulas(this);
  }

  List<ExcelFormulaCapabilityAssessment> assessAllFormulaCapabilitiesInternal() {
    return ExcelWorkbookFormulaSupport.assessAllFormulaCapabilities(this);
  }

  ExcelWorkbook evaluateFormulaCellsInternal(List<ExcelFormulaCellTarget> cells) {
    return ExcelWorkbookFormulaSupport.evaluateFormulaCells(this, cells);
  }

  List<ExcelFormulaCapabilityAssessment> assessFormulaCellCapabilitiesInternal(
      List<ExcelFormulaCellTarget> cells) {
    return ExcelWorkbookFormulaSupport.assessFormulaCellCapabilities(this, cells);
  }

  ExcelWorkbook clearFormulaCachesInternal() {
    return ExcelWorkbookFormulaSupport.clearFormulaCaches(this);
  }

  /** Resets only the in-process evaluator cache after workbook mutations. */
  ExcelWorkbook invalidateFormulaRuntime() {
    context.formulaRuntime().clearCachedResults();
    return this;
  }

  ExcelWorkbook forceFormulaRecalculationOnOpenInternal() {
    return ExcelWorkbookFormulaSupport.forceFormulaRecalculationOnOpen(this);
  }

  /** Returns the number of sheets in the workbook. */
  public int sheetCount() {
    return context.workbook().getNumberOfSheets();
  }

  /** Returns the number of analyzable named ranges currently present in the workbook. */
  public int namedRangeCount() {
    return namedRanges().size();
  }

  /** Returns an ordered list of all sheet names in the workbook. */
  public List<String> sheetNames() {
    return ExcelWorkbookSheetAccessSupport.sheetNames(this);
  }

  /** Returns every analyzable named range currently present in the workbook. */
  public List<ExcelNamedRangeSnapshot> namedRanges() {
    return ExcelWorkbookNamedRangeSupport.namedRanges(this);
  }

  boolean forceFormulaRecalculationOnOpenEnabledInternal() {
    return context.workbook().getForceFormulaRecalculation();
  }

  /** Returns the evaluator environment facts used by formula reads and diagnostics. */
  ExcelFormulaRuntimeContext formulaRuntimeContext() {
    return context.formulaRuntime().context();
  }

  /** Returns the workbook-level summary facts including active and selected sheet state. */
  WorkbookCoreResult.WorkbookSummary workbookSummary() {
    return sheetStateController().summarizeWorkbook(this);
  }

  /** Returns the workbook-level protection facts currently stored in the workbook. */
  ExcelWorkbookProtectionSnapshot workbookProtection() {
    return sheetStateController().workbookProtection(this);
  }

  /** Returns the summary facts for one sheet, including visibility and protection state. */
  WorkbookSheetResult.SheetSummary sheetSummary(String sheetName) {
    return sheetStateController().summarizeSheet(this, sheetName);
  }

  /** Saves the workbook to disk, creating parent directories as needed. */
  public void save(Path workbookPath) throws IOException {
    save(workbookPath, null);
  }

  /** Saves the workbook to disk with optional OOXML package-encryption and signing settings. */
  public void save(Path workbookPath, ExcelOoxmlPersistenceOptions persistenceOptions)
      throws IOException {
    save(workbookPath, persistenceOptions, ExcelTempFiles::createManagedTempFile);
  }

  /** Saves the workbook to disk with explicit package-security temp-file ownership. */
  public void save(
      Path workbookPath,
      ExcelOoxmlPersistenceOptions persistenceOptions,
      ExcelOoxmlPackageSecuritySupport.TempFileFactory tempFileFactory)
      throws IOException {
    ExcelWorkbookPersistenceSupport.save(this, workbookPath, persistenceOptions, tempFileFactory);
  }

  /** Saves the plain OOXML workbook package with no encryption or signing wrapper. */
  public void savePlainWorkbook(Path workbookPath) throws IOException {
    ExcelWorkbookPersistenceSupport.savePlainWorkbook(this, workbookPath);
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
  XSSFWorkbook xssfWorkbook() {
    return context.workbook();
  }

  ExcelWorkbookContext context() {
    return context;
  }

  Path sourcePath() {
    return context.sourcePath();
  }

  ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity() {
    return context.loadedPackageSecurity();
  }

  String sourceEncryptionPassword() {
    return context.sourceEncryptionPassword();
  }

  boolean wasMutatedSinceOpen() {
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

  /** Returns whether the POI defined name is a user-facing range that GridGrind should analyze. */
  static boolean shouldExpose(Name name) {
    return ExcelWorkbookNamedRangeSupport.shouldExpose(name);
  }

  /** Returns whether a defined-name triple is user-facing and analyzable by GridGrind. */
  static boolean shouldExpose(String nameName, boolean functionName, boolean hidden) {
    return ExcelWorkbookNamedRangeSupport.shouldExpose(nameName, functionName, hidden);
  }

  private static ExcelTableController tableController() {
    return new ExcelTableController();
  }

  private static ExcelPivotTableController pivotTableController() {
    return new ExcelPivotTableController();
  }

  private static ExcelCustomXmlController customXmlController() {
    return new ExcelCustomXmlController();
  }

  private static ExcelSheetCopyController sheetCopyController() {
    return new ExcelSheetCopyController();
  }

  private static ExcelSheetStateController sheetStateController() {
    return new ExcelSheetStateController();
  }
}
