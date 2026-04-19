package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * High-level workbook wrapper around Apache POI for creation, loading, saving, and sheet access.
 */
public final class ExcelWorkbook implements AutoCloseable {
  private final XSSFWorkbook workbook;
  private final WorkbookStyleRegistry styleRegistry;
  private final ExcelFormulaRuntime formulaRuntime;
  private final ExcelTableController tableController;
  private final ExcelPivotTableController pivotTableController;
  private final ExcelSheetCopyController sheetCopyController;
  private final ExcelSheetStateController sheetStateController;
  private final Path sourcePath;
  private final ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity;
  private final String sourceEncryptionPassword;
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
    this.workbook = workbook;
    this.styleRegistry = new WorkbookStyleRegistry(workbook);
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.tableController = new ExcelTableController();
    this.pivotTableController = new ExcelPivotTableController();
    this.sheetCopyController = new ExcelSheetCopyController();
    this.sheetStateController = new ExcelSheetStateController();
    this.sourcePath = sourcePath;
    this.loadedPackageSecurity =
        Objects.requireNonNull(loadedPackageSecurity, "loadedPackageSecurity must not be null");
    this.sourceEncryptionPassword = sourceEncryptionPassword;
  }

  /** Adapts a POI evaluator into the GridGrind-owned formula runtime seam. */
  ExcelWorkbook(XSSFWorkbook workbook, FormulaEvaluator formulaEvaluator) {
    this(workbook, ExcelFormulaRuntime.poi(formulaEvaluator));
  }

  /** Creates an empty XLSX workbook. */
  public static ExcelWorkbook create() {
    return new ExcelWorkbook(new XSSFWorkbook());
  }

  /** Creates an empty XLSX workbook with the supplied formula-evaluation environment. */
  public static ExcelWorkbook create(ExcelFormulaEnvironment formulaEnvironment)
      throws IOException {
    return new ExcelWorkbook(new XSSFWorkbook(), formulaEnvironment);
  }

  /** Opens an existing workbook file from disk. */
  public static ExcelWorkbook open(Path workbookPath) throws IOException {
    return open(workbookPath, (ExcelOoxmlOpenOptions) null);
  }

  /** Opens an existing workbook file from disk with optional OOXML package-open settings. */
  public static ExcelWorkbook open(Path workbookPath, ExcelOoxmlOpenOptions openOptions)
      throws IOException {
    return ExcelOoxmlPackageSecuritySupport.openWorkbook(
        workbookPath, openOptions, Files::createTempFile);
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
        workbookPath, formulaEnvironment, openOptions, Files::createTempFile);
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

  /** Returns the named sheet, creating it if necessary. */
  public ExcelSheet getOrCreateSheet(String sheetName) {
    requireSheetName(sheetName, "sheetName");

    Sheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      sheet = workbook.createSheet(sheetName);
    }

    return new ExcelSheet(sheet, styleRegistry, formulaRuntime);
  }

  /** Returns an existing sheet. */
  public ExcelSheet sheet(String sheetName) {
    requireSheetName(sheetName, "sheetName");

    return new ExcelSheet(requiredSheet(sheetName), styleRegistry, formulaRuntime);
  }

  /** Renames an existing sheet to a new destination name. */
  public ExcelWorkbook renameSheet(String sheetName, String newSheetName) {
    return sheetStateController.renameSheet(this, sheetName, newSheetName);
  }

  /** Deletes an existing sheet from the workbook. A workbook must retain at least one sheet. */
  public ExcelWorkbook deleteSheet(String sheetName) {
    return sheetStateController.deleteSheet(this, sheetName);
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  public ExcelWorkbook moveSheet(String sheetName, int targetIndex) {
    return sheetStateController.moveSheet(this, sheetName, targetIndex);
  }

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  public ExcelWorkbook copySheet(
      String sourceSheetName, String newSheetName, ExcelSheetCopyPosition position) {
    return sheetCopyController.copySheet(this, sourceSheetName, newSheetName, position);
  }

  /** Sets the active sheet and ensures it is selected. */
  public ExcelWorkbook setActiveSheet(String sheetName) {
    return sheetStateController.setActiveSheet(this, sheetName);
  }

  /** Sets the selected visible sheet set and normalizes the active tab into that selection. */
  public ExcelWorkbook setSelectedSheets(List<String> sheetNames) {
    return sheetStateController.setSelectedSheets(this, sheetNames);
  }

  /** Sets one sheet visibility while preserving a visible active selected sheet. */
  public ExcelWorkbook setSheetVisibility(String sheetName, ExcelSheetVisibility visibility) {
    return sheetStateController.setSheetVisibility(this, sheetName, visibility);
  }

  /** Enables sheet protection with the exact supported lock flags. */
  public ExcelWorkbook setSheetProtection(
      String sheetName, ExcelSheetProtectionSettings protection) {
    return setSheetProtection(sheetName, protection, null);
  }

  /** Enables sheet protection with the exact supported lock flags. */
  public ExcelWorkbook setSheetProtection(
      String sheetName, ExcelSheetProtectionSettings protection, String password) {
    return sheetStateController.setSheetProtection(this, sheetName, protection, password);
  }

  /** Disables sheet protection entirely. */
  public ExcelWorkbook clearSheetProtection(String sheetName) {
    return sheetStateController.clearSheetProtection(this, sheetName);
  }

  /** Enables workbook-level protection and password hashes with authoritative settings. */
  public ExcelWorkbook setWorkbookProtection(ExcelWorkbookProtectionSettings protection) {
    return sheetStateController.setWorkbookProtection(this, protection);
  }

  /** Clears workbook-level protection and password hashes entirely. */
  public ExcelWorkbook clearWorkbookProtection() {
    return sheetStateController.clearWorkbookProtection(this);
  }

  /** Creates or replaces one named range in workbook or sheet scope. */
  public ExcelWorkbook setNamedRange(ExcelNamedRangeDefinition definition) {
    Objects.requireNonNull(definition, "definition must not be null");

    Name name = existingName(definition.name(), definition.scope());
    if (name == null) {
      name = workbook.createName();
    }
    applyScope(name, definition.scope());
    name.setNameName(definition.name());
    name.setRefersToFormula(definition.target().refersToFormula());
    return this;
  }

  /** Deletes one named range from workbook or sheet scope. */
  public ExcelWorkbook deleteNamedRange(String name, ExcelNamedRangeScope scope) {
    Name existingName = requiredName(name, scope);
    workbook.removeName(existingName);
    return this;
  }

  /** Creates or replaces one workbook-global table definition. */
  public ExcelWorkbook setTable(ExcelTableDefinition definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    tableController.setTable(this, definition);
    return this;
  }

  /** Deletes one existing table by workbook-global name and expected sheet name. */
  public ExcelWorkbook deleteTable(String name, String sheetName) {
    tableController.deleteTable(this, name, sheetName);
    return this;
  }

  /** Creates or replaces one workbook-global pivot-table definition. */
  public ExcelWorkbook setPivotTable(ExcelPivotTableDefinition definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    pivotTableController.setPivotTable(this, definition);
    return this;
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet name. */
  public ExcelWorkbook deletePivotTable(String name, String sheetName) {
    pivotTableController.deletePivotTable(this, name, sheetName);
    return this;
  }

  /** Evaluates every formula cell currently present in the workbook. */
  public ExcelWorkbook evaluateAllFormulas() {
    formulaRuntime.clearCachedResults();
    for (Sheet sheet : workbook) {
      for (Row row : sheet) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            evaluateFormulaCell(sheet.getSheetName(), cell);
          }
        }
      }
    }
    markPackageMutated();
    return this;
  }

  /** Classifies every formula cell currently present in the workbook under the current runtime. */
  public List<ExcelFormulaCapabilityAssessment> assessAllFormulaCapabilities() {
    formulaRuntime.clearCachedResults();
    List<ExcelFormulaCapabilityAssessment> assessments = new ArrayList<>();
    for (Sheet sheet : workbook) {
      for (Row row : sheet) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            assessments.add(assessFormulaCell(sheet.getSheetName(), cell));
          }
        }
      }
    }
    formulaRuntime.clearCachedResults();
    return List.copyOf(assessments);
  }

  /** Evaluates one or more explicit formula-cell targets and stores their cached results. */
  public ExcelWorkbook evaluateFormulaCells(List<ExcelFormulaCellTarget> cells) {
    Objects.requireNonNull(cells, "cells must not be null");
    formulaRuntime.clearCachedResults();
    for (ExcelFormulaCellTarget target : cells) {
      Objects.requireNonNull(target, "cells must not contain nulls");
      Cell cell = requiredCell(target.sheetName(), target.address());
      if (cell.getCellType() != CellType.FORMULA) {
        throw new IllegalArgumentException(
            "Cell " + target.sheetName() + "!" + target.address() + " is not a formula cell");
      }
      evaluateFormulaCell(target.sheetName(), cell);
    }
    markPackageMutated();
    return this;
  }

  /** Classifies one explicit set of formula-cell targets under the current runtime. */
  public List<ExcelFormulaCapabilityAssessment> assessFormulaCellCapabilities(
      List<ExcelFormulaCellTarget> cells) {
    Objects.requireNonNull(cells, "cells must not be null");
    formulaRuntime.clearCachedResults();
    List<ExcelFormulaCapabilityAssessment> assessments = new ArrayList<>(cells.size());
    for (ExcelFormulaCellTarget target : cells) {
      Objects.requireNonNull(target, "cells must not contain nulls");
      Cell cell = requiredCell(target.sheetName(), target.address());
      if (cell.getCellType() != CellType.FORMULA) {
        throw new IllegalArgumentException(
            "Cell " + target.sheetName() + "!" + target.address() + " is not a formula cell");
      }
      assessments.add(assessFormulaCell(target.sheetName(), cell));
    }
    formulaRuntime.clearCachedResults();
    return List.copyOf(assessments);
  }

  /** Clears persisted formula cached results and resets the in-process evaluator state. */
  public ExcelWorkbook clearFormulaCaches() {
    clearPersistedFormulaCaches();
    invalidateFormulaRuntime();
    markPackageMutated();
    return this;
  }

  /** Resets only the in-process evaluator cache after workbook mutations. */
  ExcelWorkbook invalidateFormulaRuntime() {
    formulaRuntime.clearCachedResults();
    return this;
  }

  private void evaluateFormulaCell(String sheetName, Cell cell) {
    try {
      formulaRuntime.evaluateFormulaCell(cell);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(
          formulaRuntime,
          sheetName,
          new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
          cell.getCellFormula(),
          exception);
    }
  }

  private ExcelFormulaCapabilityAssessment assessFormulaCell(String sheetName, Cell cell) {
    String address = new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString();
    String formula = cell.getCellFormula();
    try {
      formulaRuntime.evaluate(cell);
      return new ExcelFormulaCapabilityAssessment(
          sheetName, address, formula, ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null);
    } catch (RuntimeException exception) {
      RuntimeException wrapped =
          FormulaExceptions.wrap(formulaRuntime, sheetName, address, formula, exception);
      return switch (wrapped) {
        case InvalidFormulaException invalid ->
            new ExcelFormulaCapabilityAssessment(
                invalid.sheetName(),
                invalid.address(),
                invalid.formula(),
                ExcelFormulaCapabilityKind.UNPARSEABLE_BY_POI,
                ExcelFormulaCapabilityIssue.INVALID_FORMULA,
                invalid.getMessage());
        case MissingExternalWorkbookException missing ->
            new ExcelFormulaCapabilityAssessment(
                missing.sheetName(),
                missing.address(),
                missing.formula(),
                ExcelFormulaCapabilityKind.UNEVALUABLE_NOW,
                ExcelFormulaCapabilityIssue.MISSING_EXTERNAL_WORKBOOK,
                missing.getMessage());
        case UnregisteredUserDefinedFunctionException unregistered ->
            new ExcelFormulaCapabilityAssessment(
                unregistered.sheetName(),
                unregistered.address(),
                unregistered.formula(),
                ExcelFormulaCapabilityKind.UNEVALUABLE_NOW,
                ExcelFormulaCapabilityIssue.UNREGISTERED_USER_DEFINED_FUNCTION,
                unregistered.getMessage());
        case UnsupportedFormulaException unsupported ->
            new ExcelFormulaCapabilityAssessment(
                unsupported.sheetName(),
                unsupported.address(),
                unsupported.formula(),
                ExcelFormulaCapabilityKind.UNEVALUABLE_NOW,
                ExcelFormulaCapabilityIssue.UNSUPPORTED_FORMULA,
                unsupported.getMessage());
        default -> throw wrapped;
      };
    }
  }

  /** Marks the workbook to recalculate formulas when opened in Excel-compatible clients. */
  public ExcelWorkbook forceFormulaRecalculationOnOpen() {
    workbook.setForceFormulaRecalculation(true);
    markPackageMutated();
    return this;
  }

  /** Returns the number of sheets in the workbook. */
  public int sheetCount() {
    return workbook.getNumberOfSheets();
  }

  /** Returns the number of analyzable named ranges currently present in the workbook. */
  public int namedRangeCount() {
    return namedRanges().size();
  }

  /** Returns an ordered list of all sheet names in the workbook. */
  public List<String> sheetNames() {
    List<String> sheetNames = new ArrayList<>(workbook.getNumberOfSheets());
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      sheetNames.add(workbook.getSheetName(sheetIndex));
    }
    return List.copyOf(sheetNames);
  }

  /** Returns every analyzable named range currently present in the workbook. */
  public List<ExcelNamedRangeSnapshot> namedRanges() {
    List<ExcelNamedRangeSnapshot> namedRanges = new ArrayList<>();
    for (Name name : workbook.getAllNames()) {
      if (!shouldExpose(name)) {
        continue;
      }
      ExcelNamedRangeScope scope = toScope(name);
      String refersToFormula = Objects.requireNonNullElse(name.getRefersToFormula(), "");
      var target = ExcelNamedRangeTargets.resolveTarget(refersToFormula, scope);
      if (target.isEmpty()) {
        namedRanges.add(
            new ExcelNamedRangeSnapshot.FormulaSnapshot(
                name.getNameName(), scope, refersToFormula));
      } else {
        namedRanges.add(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                name.getNameName(), scope, refersToFormula, target.orElseThrow()));
      }
    }
    return List.copyOf(namedRanges);
  }

  /** Returns whether the workbook is marked to recalculate formulas when opened in Excel. */
  public boolean forceFormulaRecalculationOnOpenEnabled() {
    return workbook.getForceFormulaRecalculation();
  }

  /** Returns the evaluator environment facts used by formula reads and diagnostics. */
  ExcelFormulaRuntimeContext formulaRuntimeContext() {
    return formulaRuntime.context();
  }

  /** Returns the workbook-level summary facts including active and selected sheet state. */
  WorkbookReadResult.WorkbookSummary workbookSummary() {
    return sheetStateController.summarizeWorkbook(this);
  }

  /** Returns the workbook-level protection facts currently stored in the workbook. */
  ExcelWorkbookProtectionSnapshot workbookProtection() {
    return sheetStateController.workbookProtection(this);
  }

  /** Returns the summary facts for one sheet, including visibility and protection state. */
  WorkbookReadResult.SheetSummary sheetSummary(String sheetName) {
    return sheetStateController.summarizeSheet(this, sheetName);
  }

  /** Saves the workbook to disk, creating parent directories as needed. */
  public void save(Path workbookPath) throws IOException {
    save(workbookPath, null);
  }

  /** Saves the workbook to disk with optional OOXML package-encryption and signing settings. */
  public void save(Path workbookPath, ExcelOoxmlPersistenceOptions persistenceOptions)
      throws IOException {
    save(workbookPath, persistenceOptions, Files::createTempFile);
  }

  /** Saves the workbook to disk with explicit package-security temp-file ownership. */
  public void save(
      Path workbookPath,
      ExcelOoxmlPersistenceOptions persistenceOptions,
      ExcelOoxmlPackageSecuritySupport.TempFileFactory tempFileFactory)
      throws IOException {
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    if ((persistenceOptions == null || persistenceOptions.isEmpty())
        && !loadedPackageSecurity.isSecure()) {
      savePlainWorkbook(workbookPath);
      return;
    }
    ExcelOoxmlPackageSecuritySupport.saveWorkbook(
        this, workbookPath, persistenceOptions, tempFileFactory);
  }

  /** Saves the plain OOXML workbook package with no encryption or signing wrapper. */
  public void savePlainWorkbook(Path workbookPath) throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");

    Path absolutePath = workbookPath.toAbsolutePath();
    Files.createDirectories(absolutePath.getParent());
    for (Sheet sheet : workbook) {
      ExcelRowColumnStructureController.canonicalizeColumnDefinitions(
          (org.apache.poi.xssf.usermodel.XSSFSheet) sheet);
    }
    ExcelTableHeaderSyncSupport.syncAllHeaders(workbook);

    try (OutputStream outputStream = Files.newOutputStream(absolutePath)) {
      workbook.write(outputStream);
    }
  }

  /** Returns factual OOXML package-security state for the current in-memory workbook. */
  ExcelOoxmlPackageSecuritySnapshot packageSecurity() {
    return mutatedSinceOpen ? loadedPackageSecurity.afterMutation() : loadedPackageSecurity;
  }

  @Override
  public void close() throws IOException {
    IOException failure = null;
    try {
      formulaRuntime.close();
    } catch (IOException exception) {
      failure = exception;
    }
    try {
      workbook.close();
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

  /** Returns the mutable XSSF workbook delegate used by workbook-scoped controllers. */
  XSSFWorkbook xssfWorkbook() {
    return workbook;
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

  boolean wasMutatedSinceOpen() {
    return mutatedSinceOpen;
  }

  void markPackageMutated() {
    mutatedSinceOpen = true;
  }

  private Cell requiredCell(String sheetName, String address) {
    CellReference reference = parseCellReference(address);
    Sheet sheet = requiredSheet(sheetName);
    Row row = sheet.getRow(reference.getRow());
    if (row == null) {
      throw new CellNotFoundException(address);
    }
    Cell cell = row.getCell(reference.getCol());
    if (cell == null) {
      throw new CellNotFoundException(address);
    }
    return cell;
  }

  private static CellReference parseCellReference(String address) {
    try {
      return new CellReference(address);
    } catch (IllegalArgumentException exception) {
      throw new InvalidCellAddressException(address, exception);
    }
  }

  private Sheet requiredSheet(String sheetName) {
    Sheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheet;
  }

  private int requiredSheetIndex(String sheetName) {
    requireSheetName(sheetName, "sheetName");

    int sheetIndex = workbook.getSheetIndex(sheetName);
    if (sheetIndex < 0) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheetIndex;
  }

  private void clearPersistedFormulaCaches() {
    for (Sheet sheet : workbook) {
      for (Row row : sheet) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            clearPersistedFormulaCache(cell);
          }
        }
      }
    }
  }

  private static void clearPersistedFormulaCache(Cell cell) {
    XSSFCell xssfCell = (XSSFCell) cell;
    var ctCell = xssfCell.getCTCell();
    if (ctCell.isSetV()) {
      ctCell.unsetV();
    }
    if (ctCell.isSetIs()) {
      ctCell.unsetIs();
    }
    if (ctCell.isSetT()) {
      ctCell.unsetT();
    }
  }

  private Name requiredName(String name, ExcelNamedRangeScope scope) {
    Name existingName = existingName(name, scope);
    if (existingName == null) {
      throw new NamedRangeNotFoundException(name, scope);
    }
    return existingName;
  }

  private Name existingName(String name, ExcelNamedRangeScope scope) {
    String validatedName = ExcelNamedRangeDefinition.validateName(name);
    Objects.requireNonNull(scope, "scope must not be null");

    return workbook.getAllNames().stream()
        .filter(candidate -> candidate.getNameName().equalsIgnoreCase(validatedName))
        .filter(candidate -> scopeMatches(candidate, scope))
        .findFirst()
        .orElse(null);
  }

  /** Returns whether the POI defined name belongs to the requested workbook or sheet scope. */
  boolean scopeMatches(Name candidate, ExcelNamedRangeScope scope) {
    return switch (scope) {
      case ExcelNamedRangeScope.WorkbookScope _ -> candidate.getSheetIndex() < 0;
      case ExcelNamedRangeScope.SheetScope sheetScope ->
          candidate.getSheetIndex() == requiredSheetIndex(sheetScope.sheetName());
    };
  }

  private void applyScope(Name name, ExcelNamedRangeScope scope) {
    switch (scope) {
      case ExcelNamedRangeScope.WorkbookScope _ -> name.setSheetIndex(-1);
      case ExcelNamedRangeScope.SheetScope sheetScope ->
          name.setSheetIndex(requiredSheetIndex(sheetScope.sheetName()));
    }
  }

  /** Returns whether the POI defined name is a user-facing range that GridGrind should analyze. */
  static boolean shouldExpose(Name name) {
    return shouldExpose(name.getNameName(), name.isFunctionName(), name.isHidden());
  }

  /** Returns whether a defined-name triple is user-facing and analyzable by GridGrind. */
  static boolean shouldExpose(String nameName, boolean functionName, boolean hidden) {
    return !functionName
        && !hidden
        && nameName != null
        && !nameName.startsWith("_xlnm.")
        && !nameName.startsWith("_XLNM.");
  }

  private ExcelNamedRangeScope toScope(Name name) {
    int sheetIndex = name.getSheetIndex();
    if (sheetIndex < 0) {
      return new ExcelNamedRangeScope.WorkbookScope();
    }
    return new ExcelNamedRangeScope.SheetScope(workbook.getSheetName(sheetIndex));
  }

  private static void requireSheetName(String value, String fieldName) {
    ExcelSheetNames.requireValid(value, fieldName);
  }
}
