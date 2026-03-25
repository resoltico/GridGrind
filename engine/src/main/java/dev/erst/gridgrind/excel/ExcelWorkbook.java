package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * High-level workbook wrapper around Apache POI for creation, loading, saving, and sheet access.
 */
public final class ExcelWorkbook implements AutoCloseable {
  private final Workbook workbook;
  private final WorkbookStyleRegistry styleRegistry;
  private final FormulaEvaluator formulaEvaluator;

  private ExcelWorkbook(Workbook workbook) {
    this(workbook, workbook.getCreationHelper().createFormulaEvaluator());
  }

  ExcelWorkbook(Workbook workbook, FormulaEvaluator formulaEvaluator) {
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
      return new ExcelWorkbook(WorkbookFactory.create(inputStream));
    }
  }

  /** Returns the named sheet, creating it if necessary. */
  public ExcelSheet getOrCreateSheet(String sheetName) {
    requireNonBlank(sheetName, "sheetName");

    Sheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      sheet = workbook.createSheet(sheetName);
    }

    return new ExcelSheet(sheet, styleRegistry, formulaEvaluator);
  }

  /** Returns an existing sheet. */
  public ExcelSheet sheet(String sheetName) {
    requireNonBlank(sheetName, "sheetName");

    Sheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }

    return new ExcelSheet(sheet, styleRegistry, formulaEvaluator);
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

  /** Returns an ordered list of all sheet names in the workbook. */
  public List<String> sheetNames() {
    List<String> sheetNames = new ArrayList<>(workbook.getNumberOfSheets());
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      sheetNames.add(workbook.getSheetName(sheetIndex));
    }
    return List.copyOf(sheetNames);
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
}
