package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;

/**
 * Normalizes copied or externally materialized calculated-column cells back to GridGrind's
 * metadata-owned table shape.
 */
final class ExcelTableCalculatedColumnCanonicalizer {
  private ExcelTableCalculatedColumnCanonicalizer() {}

  /**
   * Clears materialized structured-reference body formulas so POI rename/update logic only sees
   * parseable ordinary cell formulas.
   */
  static void canonicalizeWorkbook(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      canonicalizeSheet(workbook.getSheetAt(sheetIndex));
    }
  }

  /** Canonicalizes one sheet's table body cells using the authored table metadata as truth. */
  static void canonicalizeSheet(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    for (XSSFTable table : sheet.getTables()) {
      canonicalizeTable(sheet, table);
    }
  }

  private static void canonicalizeTable(XSSFSheet sheet, XSSFTable table) {
    ExcelRange range =
        ExcelSheetStructureSupport.parseRangeOrNull(
            Objects.requireNonNullElse(table.getCTTable().getRef(), ""));
    if (range == null) {
      return;
    }
    int firstBodyRow = range.firstRow() + 1;
    int lastBodyRow = range.lastRow() - totalsRowCount(table.getCTTable());
    if (firstBodyRow > lastBodyRow) {
      return;
    }

    CTTableColumn[] columns = table.getCTTable().getTableColumns().getTableColumnArray();
    for (int columnOffset = 0; columnOffset < columns.length; columnOffset++) {
      String calculatedFormula = calculatedColumnFormula(columns[columnOffset]);
      if (calculatedFormula == null) {
        continue;
      }
      int columnIndex = range.firstColumn() + columnOffset;
      for (int rowIndex = firstBodyRow; rowIndex <= lastBodyRow; rowIndex++) {
        clearMaterializedCalculatedFormulaCell(sheet, rowIndex, columnIndex, calculatedFormula);
      }
    }
  }

  private static int totalsRowCount(CTTable table) {
    return table.getTotalsRowCount() > 0 ? Math.toIntExact(table.getTotalsRowCount()) : 0;
  }

  static String calculatedColumnFormula(CTTableColumn column) {
    if (!column.isSetCalculatedColumnFormula()) {
      return null;
    }
    String formula =
        Objects.requireNonNullElse(column.getCalculatedColumnFormula().getStringValue(), "");
    return formula.isBlank() ? null : formula;
  }

  private static void clearMaterializedCalculatedFormulaCell(
      XSSFSheet sheet, int rowIndex, int columnIndex, String calculatedFormula) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      return;
    }
    Cell cell = row.getCell(columnIndex);
    if (cell == null || cell.getCellType() != CellType.FORMULA) {
      return;
    }
    if (!calculatedFormula.equals(cell.getCellFormula())) {
      return;
    }
    if (mustPreserveCellShell(cell)) {
      cell.setBlank();
      return;
    }
    row.removeCell(cell);
  }

  static boolean mustPreserveCellShell(Cell cell) {
    return cell.getCellComment() != null
        || cell.getHyperlink() != null
        || cell.getCellStyle().getIndex() != 0;
  }
}
