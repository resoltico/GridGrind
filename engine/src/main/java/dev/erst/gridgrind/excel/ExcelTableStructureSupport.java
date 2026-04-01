package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

/** Shared structural helpers for GridGrind table authoring and snapshot normalization. */
final class ExcelTableStructureSupport {
  private ExcelTableStructureSupport() {}

  /** Validates that one table range can hold its header row, data rows, and optional totals row. */
  static void requireSupportedTableShape(ExcelRange range, boolean showTotalsRow) {
    int minimumRows = showTotalsRow ? 3 : 2;
    if (range.rowCount() < minimumRows) {
      throw new IllegalArgumentException(
          "table range must contain at least "
              + minimumRows
              + " rows to cover a header row, one data row, and any totals row");
    }
  }

  /** Returns the normalized header texts contributed by one candidate header row range. */
  static List<String> headerNames(org.apache.poi.xssf.usermodel.XSSFSheet sheet, ExcelRange range) {
    Row headerRow = sheet.getRow(range.firstRow());
    List<String> headers = new ArrayList<>(range.columnCount());
    for (int columnIndex = range.firstColumn(); columnIndex <= range.lastColumn(); columnIndex++) {
      headers.add(headerText(headerRow == null ? null : headerRow.getCell(columnIndex)));
    }
    return List.copyOf(headers);
  }

  /** Returns the normalized header text contributed by one candidate header cell. */
  static String headerText(Cell cell) {
    if (cell == null) {
      return "";
    }
    return switch (cell.getCellType()) {
      case BLANK -> "";
      case _NONE -> "";
      case STRING -> cell.getStringCellValue().trim();
      case FORMULA -> cell.getCellFormula().trim();
      default -> cell.toString().trim();
    };
  }

  /** Applies one GridGrind table-owned autofilter definition to a mutable XSSF table. */
  static void applyAutofilter(XSSFTable table, ExcelRange range, boolean showTotalsRow) {
    CTAutoFilter autoFilter =
        table.getCTTable().isSetAutoFilter()
            ? table.getCTTable().getAutoFilter()
            : table.getCTTable().addNewAutoFilter();
    int lastFilterRow = showTotalsRow ? range.lastRow() - 1 : range.lastRow();
    autoFilter.setRef(
        new CellRangeAddress(
                range.firstRow(), lastFilterRow, range.firstColumn(), range.lastColumn())
            .formatAsString());
  }

  /** Applies one GridGrind table-style definition to a mutable XSSF table. */
  static void applyStyle(XSSFTable table, ExcelTableStyle style) {
    Objects.requireNonNull(table, "table must not be null");
    Objects.requireNonNull(style, "style must not be null");

    CTTable ctTable = table.getCTTable();
    switch (style) {
      case ExcelTableStyle.None _ -> {
        if (ctTable.isSetTableStyleInfo()) {
          ctTable.unsetTableStyleInfo();
        }
      }
      case ExcelTableStyle.Named named -> {
        CTTableStyleInfo styleInfo =
            ctTable.isSetTableStyleInfo()
                ? ctTable.getTableStyleInfo()
                : ctTable.addNewTableStyleInfo();
        styleInfo.setName(named.name());
        styleInfo.setShowFirstColumn(named.showFirstColumn());
        styleInfo.setShowLastColumn(named.showLastColumn());
        styleInfo.setShowRowStripes(named.showRowStripes());
        styleInfo.setShowColumnStripes(named.showColumnStripes());
      }
    }
  }

  /** Returns the table-owned autofilter range text implied by one factual table snapshot. */
  static String expectedAutofilterRangeText(ExcelTableSnapshot table) {
    Objects.requireNonNull(table, "table must not be null");

    ExcelRange range = ExcelSheetStructureSupport.parseRangeOrNull(table.range());
    if (range == null) {
      return table.range();
    }
    int lastFilterRow = range.lastRow() - Math.min(table.totalsRowCount(), range.rowCount() - 1);
    return new CellRangeAddress(
            range.firstRow(), lastFilterRow, range.firstColumn(), range.lastColumn())
        .formatAsString();
  }
}
