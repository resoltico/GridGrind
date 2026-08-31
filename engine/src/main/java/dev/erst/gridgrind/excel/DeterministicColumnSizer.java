package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentSkipListMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/** Sizes columns deterministically from displayed cell content instead of host font metrics. */
final class DeterministicColumnSizer {
  private static final double COLUMN_PADDING = 2.0d;

  private DeterministicColumnSizer() {}

  /** Applies deterministic widths to every column that currently contains visible cell content. */
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static void autoSize(Sheet sheet, DataFormatter dataFormatter) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(dataFormatter, "dataFormatter must not be null");

    Map<Integer, Double> widthsByColumn = new ConcurrentSkipListMap<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (!contributesVisibleContent(cell)) {
          continue;
        }
        double width = displayedWidthCharacters(cell, dataFormatter);
        if (width <= 0.0d) {
          continue;
        }
        widthsByColumn.merge(cell.getColumnIndex(), width, Math::max);
      }
    }

    widthsByColumn.forEach(
        (columnIndex, widthCharacters) ->
            sheet.setColumnWidth(
                columnIndex,
                ExcelSheetStructureSupport.toColumnWidthUnits(
                    Math.min(
                        widthCharacters, ExcelSheetLayoutLimits.MAX_COLUMN_WIDTH_CHARACTERS))));
  }

  private static boolean contributesVisibleContent(Cell cell) {
    return cell.getCellType() != CellType.BLANK;
  }

  private static double displayedWidthCharacters(Cell cell, DataFormatter dataFormatter) {
    String displayValue =
        cell.getCellType() == CellType.FORMULA
            ? cachedFormulaDisplayValue(cell)
            : dataFormatter.formatCellValue(cell);
    return contentWidthCharacters(displayValue);
  }

  private static String cachedFormulaDisplayValue(Cell cell) {
    DataFormatter cachedValueFormatter = new DataFormatter();
    cachedValueFormatter.setUseCachedValuesForFormulaCells(true);
    return cachedValueFormatter.formatCellValue(cell);
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
