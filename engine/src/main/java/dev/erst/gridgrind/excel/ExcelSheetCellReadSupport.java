package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Typed cell reads, snapshots, windows, and previews for one sheet wrapper. */
final class ExcelSheetCellReadSupport {
  private final Sheet sheet;
  private final WorkbookStyleRegistry styleRegistry;
  private final ExcelFormulaRuntime formulaRuntime;
  private final DataFormatter dataFormatter;
  private final ExcelSheetAnnotationSupport annotationSupport;

  ExcelSheetCellReadSupport(
      Sheet sheet,
      WorkbookStyleRegistry styleRegistry,
      ExcelFormulaRuntime formulaRuntime,
      DataFormatter dataFormatter,
      ExcelSheetAnnotationSupport annotationSupport) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.styleRegistry = Objects.requireNonNull(styleRegistry, "styleRegistry must not be null");
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.dataFormatter = Objects.requireNonNull(dataFormatter, "dataFormatter must not be null");
    this.annotationSupport =
        Objects.requireNonNull(annotationSupport, "annotationSupport must not be null");
  }

  String text(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    return requiredCell(address).getStringCellValue();
  }

  double number(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() == CellType.FORMULA) {
      CellValue evaluatedCell;
      try {
        evaluatedCell = formulaRuntime.evaluate(cell);
      } catch (RuntimeException exception) {
        throw FormulaExceptions.wrap(
            formulaRuntime, sheet.getSheetName(), address, cell.getCellFormula(), exception);
      }
      if (evaluatedCell == null || evaluatedCell.getCellType() != CellType.NUMERIC) {
        throw new IllegalStateException("Cell does not evaluate to a numeric value: " + address);
      }
      return evaluatedCell.getNumberValue();
    }
    return cell.getNumericCellValue();
  }

  boolean bool(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() == CellType.FORMULA) {
      CellValue evaluatedCell;
      try {
        evaluatedCell = formulaRuntime.evaluate(cell);
      } catch (RuntimeException exception) {
        throw FormulaExceptions.wrap(
            formulaRuntime, sheet.getSheetName(), address, cell.getCellFormula(), exception);
      }
      if (evaluatedCell == null || evaluatedCell.getCellType() != CellType.BOOLEAN) {
        throw new IllegalStateException("Cell does not evaluate to a boolean value: " + address);
      }
      return evaluatedCell.getBooleanValue();
    }
    return cell.getBooleanCellValue();
  }

  String formula(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() != CellType.FORMULA) {
      throw new IllegalStateException("Cell does not contain a formula: " + address);
    }
    return cell.getCellFormula();
  }

  ExcelCellSnapshot snapshotCell(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    CellReference cellReference = parseCellReference(address);
    requireValidCellReference(address, cellReference);
    return snapshotCellOrBlank(address, cellReference.getRow(), cellReference.getCol());
  }

  List<ExcelCellSnapshot> snapshotCells(List<String> addresses) {
    Objects.requireNonNull(addresses, "addresses must not be null");
    List<ExcelCellSnapshot> cells = new ArrayList<>(addresses.size());
    for (String address : List.copyOf(addresses)) {
      ExcelSheet.requireNonBlank(address, "address");
      cells.add(snapshotCell(address));
    }
    return List.copyOf(cells);
  }

  List<ExcelPreviewRow> preview(int maxRows, int maxColumns, int lastRowIndex) {
    if (maxRows <= 0) {
      throw new IllegalArgumentException("maxRows must be greater than 0");
    }
    if (maxColumns <= 0) {
      throw new IllegalArgumentException("maxColumns must be greater than 0");
    }

    List<ExcelPreviewRow> previewRows = new ArrayList<>();
    int endingRowIndex = Math.min(lastRowIndex, maxRows - 1);
    for (int rowIndex = 0; rowIndex <= endingRowIndex; rowIndex++) {
      previewRows.add(new ExcelPreviewRow(rowIndex, previewRow(rowIndex, maxColumns)));
    }
    return List.copyOf(previewRows);
  }

  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  WorkbookReadResult.Window window(String topLeftAddress, int rowCount, int columnCount) {
    ExcelSheet.requireNonBlank(topLeftAddress, "topLeftAddress");
    if (rowCount <= 0) {
      throw new IllegalArgumentException("rowCount must be greater than 0");
    }
    if (columnCount <= 0) {
      throw new IllegalArgumentException("columnCount must be greater than 0");
    }

    CellReference topLeft = parseCellReference(topLeftAddress);
    requireValidCellReference(topLeftAddress, topLeft);
    int lastRow = topLeft.getRow() + rowCount - 1;
    int lastCol = topLeft.getCol() + columnCount - 1;
    if (lastRow > SpreadsheetVersion.EXCEL2007.getLastRowIndex()
        || lastCol > SpreadsheetVersion.EXCEL2007.getLastColumnIndex()) {
      throw new IllegalArgumentException(
          "window extends beyond the Excel sheet boundary: topLeft="
              + topLeftAddress
              + " rowCount="
              + rowCount
              + " columnCount="
              + columnCount);
    }
    List<WorkbookReadResult.WindowRow> rows = new ArrayList<>(rowCount);
    for (int rowOffset = 0; rowOffset < rowCount; rowOffset++) {
      int rowIndex = topLeft.getRow() + rowOffset;
      List<ExcelCellSnapshot> cells = new ArrayList<>(columnCount);
      for (int columnOffset = 0; columnOffset < columnCount; columnOffset++) {
        int columnIndex = topLeft.getCol() + columnOffset;
        String address = new CellReference(rowIndex, columnIndex).formatAsString();
        cells.add(snapshotCellOrBlank(address, rowIndex, columnIndex));
      }
      rows.add(new WorkbookReadResult.WindowRow(rowIndex, List.copyOf(cells)));
    }
    return new WorkbookReadResult.Window(
        sheet.getSheetName(), topLeftAddress, rowCount, columnCount, List.copyOf(rows));
  }

  List<ExcelCellSnapshot.FormulaSnapshot> formulaCells() {
    List<ExcelCellSnapshot.FormulaSnapshot> formulas = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() == CellType.FORMULA) {
          formulas.add(
              (ExcelCellSnapshot.FormulaSnapshot)
                  snapshot(
                      new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
                      cell));
        }
      }
    }
    return List.copyOf(formulas);
  }

  private Cell requiredCell(String address) {
    CellReference cellReference = parseCellReference(address);
    Row row = sheet.getRow(cellReference.getRow());
    if (row == null) {
      throw new CellNotFoundException(address);
    }

    Cell cell = row.getCell(cellReference.getCol());
    if (cell == null) {
      throw new CellNotFoundException(address);
    }

    return cell;
  }

  private List<ExcelCellSnapshot> previewRow(int rowIndex, int maxColumns) {
    Row row = sheet.getRow(rowIndex);
    List<ExcelCellSnapshot> cells = new ArrayList<>();
    if (row != null) {
      for (int columnIndex = 0; columnIndex < maxColumns; columnIndex++) {
        Cell cell = row.getCell(columnIndex);
        if (ExcelSheetStructureSupport.shouldPreview(cell)) {
          cells.add(snapshot(new CellReference(rowIndex, columnIndex).formatAsString(), cell));
        }
      }
    }
    return List.copyOf(cells);
  }

  private ExcelCellSnapshot snapshotCellOrBlank(String address, int rowIndex, int columnIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      return blankSnapshot(address);
    }
    Cell cell = row.getCell(columnIndex);
    if (cell == null) {
      return blankSnapshot(address);
    }
    return snapshot(address, cell);
  }

  private ExcelCellSnapshot snapshot(String address, Cell cell) {
    CellType declaredType = cell.getCellType();
    String formulaExpression = declaredType == CellType.FORMULA ? cell.getCellFormula() : null;
    String displayValue;
    try {
      displayValue =
          declaredType == CellType.FORMULA
              ? formulaRuntime.displayValue(dataFormatter, cell)
              : dataFormatter.formatCellValue(cell);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(
          formulaRuntime, sheet.getSheetName(), address, formulaExpression, exception);
    }
    ExcelCellStyleSnapshot style = styleRegistry.snapshot(cell);
    ExcelCellMetadataSnapshot metadata = annotationSupport.metadata(cell);

    if (declaredType == CellType.FORMULA) {
      String formula = cell.getCellFormula();
      CellValue evaluatedCell;
      try {
        evaluatedCell = formulaRuntime.evaluate(cell);
      } catch (RuntimeException exception) {
        throw FormulaExceptions.wrap(
            formulaRuntime, sheet.getSheetName(), address, formula, exception);
      }

      CellType evalType = evaluatedCell != null ? evaluatedCell.getCellType() : CellType.BLANK;
      ExcelCellSnapshot evaluation =
          switch (evalType) {
            case STRING ->
                new ExcelCellSnapshot.TextSnapshot(
                    address,
                    "STRING",
                    displayValue,
                    style,
                    metadata,
                    evaluatedCell.getStringValue(),
                    null);
            case NUMERIC ->
                new ExcelCellSnapshot.NumberSnapshot(
                    address,
                    "NUMBER",
                    displayValue,
                    style,
                    metadata,
                    evaluatedCell.getNumberValue());
            case BOOLEAN ->
                new ExcelCellSnapshot.BooleanSnapshot(
                    address,
                    "BOOLEAN",
                    displayValue,
                    style,
                    metadata,
                    evaluatedCell.getBooleanValue());
            case ERROR ->
                new ExcelCellSnapshot.ErrorSnapshot(
                    address,
                    "ERROR",
                    displayValue,
                    style,
                    metadata,
                    FormulaError.forInt(evaluatedCell.getErrorValue()).getString());
            default ->
                new ExcelCellSnapshot.BlankSnapshot(
                    address, "BLANK", displayValue, style, metadata);
          };
      return new ExcelCellSnapshot.FormulaSnapshot(
          address, "FORMULA", displayValue, style, metadata, formula, evaluation);
    }

    return switch (declaredType) {
      case STRING ->
          new ExcelCellSnapshot.TextSnapshot(
              address,
              "STRING",
              displayValue,
              style,
              metadata,
              cell.getStringCellValue(),
              ExcelRichTextSupport.snapshot(
                  xssfSheet().getWorkbook(),
                  (XSSFRichTextString) cell.getRichStringCellValue(),
                  style.font()));
      case NUMERIC ->
          new ExcelCellSnapshot.NumberSnapshot(
              address, "NUMBER", displayValue, style, metadata, cell.getNumericCellValue());
      case BOOLEAN ->
          new ExcelCellSnapshot.BooleanSnapshot(
              address, "BOOLEAN", displayValue, style, metadata, cell.getBooleanCellValue());
      case ERROR ->
          new ExcelCellSnapshot.ErrorSnapshot(
              address,
              "ERROR",
              displayValue,
              style,
              metadata,
              FormulaError.forInt(cell.getErrorCellValue()).getString());
      case BLANK, _NONE, FORMULA ->
          new ExcelCellSnapshot.BlankSnapshot(address, "BLANK", displayValue, style, metadata);
    };
  }

  private ExcelCellSnapshot blankSnapshot(String address) {
    return new ExcelCellSnapshot.BlankSnapshot(
        address, "BLANK", "", styleRegistry.defaultSnapshot(), ExcelCellMetadataSnapshot.empty());
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  private static CellReference parseCellReference(String address) {
    try {
      return new CellReference(address);
    } catch (IllegalArgumentException exception) {
      throw new InvalidCellAddressException(address, exception);
    }
  }

  private static void requireValidCellReference(String address, CellReference cellReference) {
    int row = cellReference.getRow();
    int col = cellReference.getCol();
    if (row < 0
        || col < 0
        || row > SpreadsheetVersion.EXCEL2007.getLastRowIndex()
        || col > SpreadsheetVersion.EXCEL2007.getLastColumnIndex()) {
      throw new InvalidCellAddressException(
          address, new IllegalArgumentException("not a valid A1-style cell address: " + address));
    }
  }
}
