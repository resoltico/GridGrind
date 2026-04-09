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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * High-level workbook wrapper around Apache POI for creation, loading, saving, and sheet access.
 */
public final class ExcelWorkbook implements AutoCloseable {
  private final XSSFWorkbook workbook;
  private final WorkbookStyleRegistry styleRegistry;
  private final ExcelFormulaRuntime formulaRuntime;
  private final ExcelTableController tableController;
  private final ExcelSheetCopyController sheetCopyController;
  private final ExcelSheetStateController sheetStateController;

  private ExcelWorkbook(XSSFWorkbook workbook) {
    this(workbook, ExcelFormulaRuntime.poi(workbook.getCreationHelper().createFormulaEvaluator()));
  }

  ExcelWorkbook(XSSFWorkbook workbook, ExcelFormulaRuntime formulaRuntime) {
    this.workbook = workbook;
    this.styleRegistry = new WorkbookStyleRegistry(workbook);
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.tableController = new ExcelTableController();
    this.sheetCopyController = new ExcelSheetCopyController();
    this.sheetStateController = new ExcelSheetStateController();
  }

  /** Adapts a POI evaluator into the GridGrind-owned formula runtime seam. */
  ExcelWorkbook(XSSFWorkbook workbook, FormulaEvaluator formulaEvaluator) {
    this(workbook, ExcelFormulaRuntime.poi(formulaEvaluator));
  }

  /** Creates an empty XLSX workbook. */
  public static ExcelWorkbook create() {
    return new ExcelWorkbook(new XSSFWorkbook());
  }

  /** Opens an existing workbook file from disk. */
  public static ExcelWorkbook open(Path workbookPath) throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");

    Path absolutePath = workbookPath.toAbsolutePath();
    if (!Files.exists(absolutePath)) {
      throw new WorkbookNotFoundException(absolutePath);
    }

    try (InputStream inputStream = Files.newInputStream(absolutePath)) {
      try {
        return new ExcelWorkbook(new XSSFWorkbook(inputStream));
      } catch (NotOfficeXmlFileException exception) {
        throw new IllegalArgumentException("Only .xlsx workbooks are supported", exception);
      }
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
    return sheetStateController.setSheetProtection(this, sheetName, protection);
  }

  /** Disables sheet protection entirely. */
  public ExcelWorkbook clearSheetProtection(String sheetName) {
    return sheetStateController.clearSheetProtection(this, sheetName);
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

  /** Evaluates every formula cell currently present in the workbook. */
  public ExcelWorkbook evaluateAllFormulas() {
    for (Sheet sheet : workbook) {
      for (Row row : sheet) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            evaluateFormulaCell(sheet.getSheetName(), cell);
          }
        }
      }
    }
    return this;
  }

  private void evaluateFormulaCell(String sheetName, Cell cell) {
    try {
      formulaRuntime.evaluate(cell);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(
          sheetName,
          new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
          cell.getCellFormula(),
          exception);
    }
  }

  /** Marks the workbook to recalculate formulas when opened in Excel-compatible clients. */
  public ExcelWorkbook forceFormulaRecalculationOnOpen() {
    workbook.setForceFormulaRecalculation(true);
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

  /** Returns the workbook-level summary facts including active and selected sheet state. */
  WorkbookReadResult.WorkbookSummary workbookSummary() {
    return sheetStateController.summarizeWorkbook(this);
  }

  /** Returns the summary facts for one sheet, including visibility and protection state. */
  WorkbookReadResult.SheetSummary sheetSummary(String sheetName) {
    return sheetStateController.summarizeSheet(this, sheetName);
  }

  /** Saves the workbook to disk, creating parent directories as needed. */
  public void save(Path workbookPath) throws IOException {
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

  @Override
  public void close() throws IOException {
    workbook.close();
  }

  /** Returns the mutable XSSF workbook delegate used by workbook-scoped controllers. */
  XSSFWorkbook xssfWorkbook() {
    return workbook;
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
