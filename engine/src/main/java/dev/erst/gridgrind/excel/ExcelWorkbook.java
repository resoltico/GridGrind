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
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * High-level workbook wrapper around Apache POI for creation, loading, saving, and sheet access.
 */
public final class ExcelWorkbook implements AutoCloseable {
  private final XSSFWorkbook workbook;
  private final WorkbookStyleRegistry styleRegistry;
  private final FormulaEvaluator formulaEvaluator;

  private ExcelWorkbook(XSSFWorkbook workbook) {
    this(workbook, workbook.getCreationHelper().createFormulaEvaluator());
  }

  ExcelWorkbook(XSSFWorkbook workbook, FormulaEvaluator formulaEvaluator) {
    this.workbook = workbook;
    this.styleRegistry = new WorkbookStyleRegistry(workbook);
    this.formulaEvaluator =
        Objects.requireNonNull(formulaEvaluator, "formulaEvaluator must not be null");
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

    return new ExcelSheet(sheet, styleRegistry, formulaEvaluator);
  }

  /** Returns an existing sheet. */
  public ExcelSheet sheet(String sheetName) {
    requireSheetName(sheetName, "sheetName");

    return new ExcelSheet(requiredSheet(sheetName), styleRegistry, formulaEvaluator);
  }

  /** Renames an existing sheet to a new destination name. */
  public ExcelWorkbook renameSheet(String sheetName, String newSheetName) {
    requireNonBlank(sheetName, "sheetName");
    requireNonBlank(newSheetName, "newSheetName");

    int sheetIndex = requiredSheetIndex(sheetName);
    WorkbookUtil.validateSheetName(newSheetName);
    requireSheetNameAvailable(newSheetName, sheetIndex);
    if (sheetName.equals(newSheetName)) {
      return this;
    }
    workbook.setSheetName(sheetIndex, newSheetName);
    return this;
  }

  /** Deletes an existing sheet from the workbook. */
  public ExcelWorkbook deleteSheet(String sheetName) {
    workbook.removeSheetAt(requiredSheetIndex(sheetName));
    return this;
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  public ExcelWorkbook moveSheet(String sheetName, int targetIndex) {
    requireNonBlank(sheetName, "sheetName");
    Sheet sheet = requiredSheet(sheetName);
    requireTargetIndex(targetIndex);

    workbook.setSheetOrder(sheet.getSheetName(), targetIndex);
    return this;
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
      formulaEvaluator.evaluate(cell);
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

  /** Saves the workbook to disk, creating parent directories as needed. */
  public void save(Path workbookPath) throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");

    Path absolutePath = workbookPath.toAbsolutePath();
    Files.createDirectories(absolutePath.getParent());

    try (OutputStream outputStream = Files.newOutputStream(absolutePath)) {
      workbook.write(outputStream);
    }
  }

  @Override
  public void close() throws IOException {
    workbook.close();
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }

  private static void requireSheetName(String value, String fieldName) {
    requireNonBlank(value, fieldName);
    if (value.length() > 31) {
      throw new IllegalArgumentException(fieldName + " must not exceed 31 characters: " + value);
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
    requireNonBlank(sheetName, "sheetName");

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
    String nameName = name.getNameName();
    return !name.isFunctionName()
        && !name.isHidden()
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

  private void requireSheetNameAvailable(String newSheetName, int currentSheetIndex) {
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      if (sheetIndex == currentSheetIndex) {
        continue;
      }
      if (workbook.getSheetName(sheetIndex).equalsIgnoreCase(newSheetName)) {
        throw new IllegalArgumentException("Sheet already exists: " + newSheetName);
      }
    }
  }

  private void requireTargetIndex(int targetIndex) {
    int sheetCount = workbook.getNumberOfSheets();
    if (targetIndex < 0 || targetIndex >= sheetCount) {
      throw new IllegalArgumentException(
          "targetIndex out of range: workbook has "
              + sheetCount
              + " sheet(s), valid positions are 0 to "
              + (sheetCount - 1)
              + "; got "
              + targetIndex);
    }
  }
}
