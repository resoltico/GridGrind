package dev.erst.gridgrind.excel;

import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentSkipListMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

/** Sizes columns deterministically from displayed cell content instead of host font metrics. */
final class DeterministicColumnSizer {
  private static final double COLUMN_PADDING = 2.0d;

  private DeterministicColumnSizer() {}

  /** Applies deterministic widths to every column that currently contains visible cell content. */
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static void autoSize(
      Sheet sheet,
      String sheetName,
      DataFormatter dataFormatter,
      ExcelFormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(dataFormatter, "dataFormatter must not be null");
    Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");

    Map<Integer, Double> widthsByColumn = new ConcurrentSkipListMap<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (!contributesVisibleContent(cell)) {
          continue;
        }
        double width = displayedWidthCharacters(sheetName, cell, dataFormatter, formulaRuntime);
        if (width <= 0.0d) {
          continue;
        }
        widthsByColumn.merge(cell.getColumnIndex(), width, Math::max);
      }
    }

    widthsByColumn.forEach(
        (columnIndex, widthCharacters) ->
            sheet.setColumnWidth(
                columnIndex, ExcelSheet.toColumnWidthUnits(Math.min(widthCharacters, 255.0d))));
  }

  private static boolean contributesVisibleContent(Cell cell) {
    return cell.getCellType() != CellType.BLANK;
  }

  private static double displayedWidthCharacters(
      String sheetName,
      Cell cell,
      DataFormatter dataFormatter,
      ExcelFormulaRuntime formulaRuntime) {
    String displayValue;
    try {
      displayValue =
          cell.getCellType() == CellType.FORMULA
              ? formulaRuntime.displayValue(dataFormatter, cell)
              : dataFormatter.formatCellValue(cell);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(
          formulaRuntime,
          sheetName,
          new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
          cell.getCellType() == CellType.FORMULA ? cell.getCellFormula() : null,
          exception);
    }
    return contentWidthCharacters(displayValue);
  }

  /** Returns the deterministic character width for one display string, including column padding. */
  static double contentWidthCharacters(String displayValue) {
    Objects.requireNonNull(displayValue, "displayValue must not be null");
    if (displayValue.isEmpty()) {
      return 0.0d;
    }
    double widestLine = 0.0d;
    for (String line : displayValue.split("\\R", -1)) {
      widestLine = Math.max(widestLine, line.codePointCount(0, line.length()) + COLUMN_PADDING);
    }
    return widestLine;
  }
}
