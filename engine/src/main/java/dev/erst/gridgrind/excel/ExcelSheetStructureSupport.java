package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelIndexDisplay;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Row, column, merge, view, and print-layout operations for one sheet wrapper. */
final class ExcelSheetStructureSupport {
  private final Sheet sheet;
  private final ExcelFormulaRuntime formulaRuntime;
  private final DataFormatter dataFormatter;
  private final ExcelPrintLayoutController printLayoutController;
  private final ExcelSheetPresentationController sheetPresentationController;
  private final ExcelRowColumnStructureController rowColumnStructureController;

  ExcelSheetStructureSupport(Sheet sheet, ExcelFormulaRuntime formulaRuntime) {
    this(
        sheet,
        formulaRuntime,
        new DataFormatter(),
        new ExcelPrintLayoutController(),
        new ExcelSheetPresentationController(),
        new ExcelRowColumnStructureController());
  }

  ExcelSheetStructureSupport(
      Sheet sheet,
      ExcelFormulaRuntime formulaRuntime,
      DataFormatter dataFormatter,
      ExcelPrintLayoutController printLayoutController,
      ExcelSheetPresentationController sheetPresentationController,
      ExcelRowColumnStructureController rowColumnStructureController) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.dataFormatter = Objects.requireNonNull(dataFormatter, "dataFormatter must not be null");
    this.printLayoutController =
        Objects.requireNonNull(printLayoutController, "printLayoutController must not be null");
    this.sheetPresentationController =
        Objects.requireNonNull(
            sheetPresentationController, "sheetPresentationController must not be null");
    this.rowColumnStructureController =
        Objects.requireNonNull(
            rowColumnStructureController, "rowColumnStructureController must not be null");
  }

  ExcelSheet mergeCells(String range, ExcelSheet owner) {
    requireNonBlank(range, "range");
    ExcelRange excelRange = ExcelRange.parse(range);
    requireMergeableRange(range, excelRange);
    if (findMergedRegionIndex(sheet, excelRange) >= 0) {
      return owner;
    }
    requireNoMergedRegionOverlap(sheet, excelRange);
    sheet.addMergedRegion(toCellRangeAddress(excelRange));
    return owner;
  }

  ExcelSheet unmergeCells(String range, ExcelSheet owner) {
    requireNonBlank(range, "range");
    ExcelRange excelRange = ExcelRange.parse(range);
    int mergedRegionIndex = findMergedRegionIndex(sheet, excelRange);
    if (mergedRegionIndex < 0) {
      throw new IllegalArgumentException("No merged region matches range: " + range);
    }
    sheet.removeMergedRegion(mergedRegionIndex);
    return owner;
  }

  ExcelSheet setColumnWidth(
      int firstColumnIndex, int lastColumnIndex, double widthCharacters, ExcelSheet owner) {
    requireNonNegative(firstColumnIndex, "firstColumnIndex");
    requireNonNegative(lastColumnIndex, "lastColumnIndex");
    requireOrderedSpan(firstColumnIndex, lastColumnIndex, "firstColumnIndex", "lastColumnIndex");

    int widthUnits = toColumnWidthUnits(widthCharacters);
    for (int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++) {
      sheet.setColumnWidth(columnIndex, widthUnits);
    }
    return owner;
  }

  ExcelSheet setRowHeight(
      int firstRowIndex, int lastRowIndex, double heightPoints, ExcelSheet owner) {
    requireNonNegative(firstRowIndex, "firstRowIndex");
    requireNonNegative(lastRowIndex, "lastRowIndex");
    requireOrderedSpan(firstRowIndex, lastRowIndex, "firstRowIndex", "lastRowIndex");

    float heightPointsValue = toRowHeightPoints(heightPoints);
    for (int rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++) {
      getOrCreateRow(rowIndex).setHeightInPoints(heightPointsValue);
    }
    return owner;
  }

  ExcelSheet insertRows(int rowIndex, int rowCount, ExcelSheet owner) {
    rowColumnStructureController.insertRows(xssfSheet(), rowIndex, rowCount);
    return owner;
  }

  ExcelSheet deleteRows(ExcelRowSpan rows, ExcelSheet owner) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowColumnStructureController.deleteRows(xssfSheet(), rows);
    return owner;
  }

  ExcelSheet shiftRows(ExcelRowSpan rows, int delta, ExcelSheet owner) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowColumnStructureController.shiftRows(xssfSheet(), rows, delta);
    return owner;
  }

  ExcelSheet insertColumns(int columnIndex, int columnCount, ExcelSheet owner) {
    rowColumnStructureController.insertColumns(xssfSheet(), columnIndex, columnCount);
    return owner;
  }

  ExcelSheet deleteColumns(ExcelColumnSpan columns, ExcelSheet owner) {
    Objects.requireNonNull(columns, "columns must not be null");
    rowColumnStructureController.deleteColumns(xssfSheet(), columns);
    return owner;
  }

  ExcelSheet shiftColumns(ExcelColumnSpan columns, int delta, ExcelSheet owner) {
    Objects.requireNonNull(columns, "columns must not be null");
    rowColumnStructureController.shiftColumns(xssfSheet(), columns, delta);
    return owner;
  }

  ExcelSheet setRowVisibility(ExcelRowSpan rows, boolean hidden, ExcelSheet owner) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowColumnStructureController.setRowVisibility(xssfSheet(), rows, hidden);
    return owner;
  }

  ExcelSheet setColumnVisibility(ExcelColumnSpan columns, boolean hidden, ExcelSheet owner) {
    Objects.requireNonNull(columns, "columns must not be null");
    rowColumnStructureController.setColumnVisibility(xssfSheet(), columns, hidden);
    return owner;
  }

  ExcelSheet groupRows(ExcelRowSpan rows, boolean collapsed, ExcelSheet owner) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowColumnStructureController.groupRows(xssfSheet(), rows, collapsed);
    return owner;
  }

  ExcelSheet ungroupRows(ExcelRowSpan rows, ExcelSheet owner) {
    Objects.requireNonNull(rows, "rows must not be null");
    rowColumnStructureController.ungroupRows(xssfSheet(), rows);
    return owner;
  }

  ExcelSheet groupColumns(ExcelColumnSpan columns, boolean collapsed, ExcelSheet owner) {
    Objects.requireNonNull(columns, "columns must not be null");
    rowColumnStructureController.groupColumns(xssfSheet(), columns, collapsed);
    return owner;
  }

  ExcelSheet ungroupColumns(ExcelColumnSpan columns, ExcelSheet owner) {
    Objects.requireNonNull(columns, "columns must not be null");
    rowColumnStructureController.ungroupColumns(xssfSheet(), columns);
    return owner;
  }

  ExcelSheet setPane(ExcelSheetPane pane, ExcelSheet owner) {
    Objects.requireNonNull(pane, "pane must not be null");
    ExcelSheetViewSupport.setPane(xssfSheet(), pane);
    return owner;
  }

  ExcelSheet setZoom(int zoomPercent, ExcelSheet owner) {
    ExcelSheetViewSupport.requireZoomPercent(zoomPercent);
    ExcelSheetViewSupport.setZoomPercent(xssfSheet(), zoomPercent);
    return owner;
  }

  ExcelSheet setPresentation(ExcelSheetPresentation presentation, ExcelSheet owner) {
    Objects.requireNonNull(presentation, "presentation must not be null");
    sheetPresentationController.setPresentation(xssfSheet(), presentation);
    return owner;
  }

  ExcelSheet setPrintLayout(ExcelPrintLayout printLayout, ExcelSheet owner) {
    Objects.requireNonNull(printLayout, "printLayout must not be null");
    printLayoutController.setPrintLayout(xssfSheet(), printLayout);
    return owner;
  }

  ExcelSheet clearPrintLayout(ExcelSheet owner) {
    printLayoutController.clearPrintLayout(xssfSheet());
    return owner;
  }

  int physicalRowCount() {
    return sheet.getPhysicalNumberOfRows();
  }

  int lastRowIndex() {
    return sheet.getLastRowNum();
  }

  int lastColumnIndex() {
    return rowColumnStructureController.lastColumnIndex(xssfSheet());
  }

  ExcelSheet autoSizeColumns(ExcelSheet owner, String sheetName) {
    DeterministicColumnSizer.autoSize(sheet, sheetName, dataFormatter, formulaRuntime);
    return owner;
  }

  List<WorkbookReadResult.MergedRegion> mergedRegions() {
    List<WorkbookReadResult.MergedRegion> mergedRegions =
        new ArrayList<>(sheet.getNumMergedRegions());
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      mergedRegions.add(
          new WorkbookReadResult.MergedRegion(sheet.getMergedRegion(regionIndex).formatAsString()));
    }
    return List.copyOf(mergedRegions);
  }

  WorkbookReadResult.SheetLayout layout(String sheetName) {
    return new WorkbookReadResult.SheetLayout(
        sheetName,
        ExcelSheetViewSupport.pane(xssfSheet()),
        ExcelSheetViewSupport.zoomPercent(xssfSheet()),
        sheetPresentationController.presentation(xssfSheet()),
        rowColumnStructureController.columnLayouts(xssfSheet()),
        rowColumnStructureController.rowLayouts(xssfSheet()));
  }

  ExcelPrintLayout printLayout() {
    return printLayoutController.printLayout(xssfSheet());
  }

  ExcelPrintLayoutSnapshot printLayoutSnapshot() {
    return printLayoutController.printLayoutSnapshot(xssfSheet());
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  private Row getOrCreateRow(int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    return row;
  }

  private static void requireMergeableRange(String range, ExcelRange excelRange) {
    if (excelRange.rowCount() == 1 && excelRange.columnCount() == 1) {
      throw new IllegalArgumentException("range must span at least two cells: " + range);
    }
  }

  static void requireNoMergedRegionOverlap(Sheet sheet, ExcelRange excelRange) {
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      CellRangeAddress existing = sheet.getMergedRegion(regionIndex);
      if (intersects(existing, excelRange)) {
        throw new IllegalArgumentException(
            "Merged range overlaps existing merged region: " + existing.formatAsString());
      }
    }
  }

  static int findMergedRegionIndex(Sheet sheet, ExcelRange excelRange) {
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      if (matches(sheet.getMergedRegion(regionIndex), excelRange)) {
        return regionIndex;
      }
    }
    return -1;
  }

  static CellRangeAddress toCellRangeAddress(ExcelRange excelRange) {
    return new CellRangeAddress(
        excelRange.firstRow(),
        excelRange.lastRow(),
        excelRange.firstColumn(),
        excelRange.lastColumn());
  }

  static ExcelRange parseRangeOrNull(String range) {
    try {
      return ExcelRange.parse(range);
    } catch (IllegalArgumentException exception) {
      return java.util.Optional.<ExcelRange>empty().orElse(null);
    }
  }

  static String formatRange(ExcelRange range) {
    return toCellRangeAddress(range).formatAsString();
  }

  static boolean matches(CellRangeAddress rangeAddress, ExcelRange excelRange) {
    return rangeAddress.getFirstRow() == excelRange.firstRow()
        && rangeAddress.getLastRow() == excelRange.lastRow()
        && rangeAddress.getFirstColumn() == excelRange.firstColumn()
        && rangeAddress.getLastColumn() == excelRange.lastColumn();
  }

  static boolean intersects(CellRangeAddress rangeAddress, ExcelRange excelRange) {
    return rangeAddress.getFirstRow() <= excelRange.lastRow()
        && rangeAddress.getLastRow() >= excelRange.firstRow()
        && rangeAddress.getFirstColumn() <= excelRange.lastColumn()
        && rangeAddress.getLastColumn() >= excelRange.firstColumn();
  }

  static boolean intersects(ExcelRange first, ExcelRange second) {
    return first.firstRow() <= second.lastRow()
        && first.lastRow() >= second.firstRow()
        && first.firstColumn() <= second.lastColumn()
        && first.lastColumn() >= second.firstColumn();
  }

  static boolean hasHeaderValue(Cell cell) {
    return !ExcelTableStructureSupport.headerText(cell).isBlank();
  }

  static boolean headerRowMissing(XSSFSheet sheet, ExcelRange range) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(range, "range must not be null");
    Row headerRow = sheet.getRow(range.firstRow());
    if (headerRow == null) {
      return true;
    }
    for (int columnIndex = range.firstColumn(); columnIndex <= range.lastColumn(); columnIndex++) {
      if (hasHeaderValue(headerRow.getCell(columnIndex))) {
        return false;
      }
    }
    return true;
  }

  static int toColumnWidthUnits(double widthCharacters) {
    ExcelSheetLayoutLimits.requireColumnWidthCharacters(widthCharacters, "widthCharacters");
    int widthUnits = (int) Math.round(widthCharacters * 256.0d);
    return widthUnits;
  }

  static float toRowHeightPoints(double heightPoints) {
    ExcelSheetLayoutLimits.requireRowHeightPoints(heightPoints, "heightPoints");
    return (float) heightPoints;
  }

  static boolean shouldPreview(Cell cell) {
    return cell != null
        && shouldPreview(
            cell.getCellType(),
            cell.getCellStyle().getIndex(),
            cell.getHyperlink() != null,
            cell.getCellComment() != null);
  }

  static boolean shouldPreview(
      CellType cellType, short styleIndex, boolean hasHyperlink, boolean hasComment) {
    return cellType != CellType.BLANK || styleIndex != 0 || hasHyperlink || hasComment;
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }

  private static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(ExcelIndexDisplay.mustNotBeNegative(fieldName, value));
    }
  }

  private static void requireOrderedSpan(
      int firstValue, int lastValue, String firstFieldName, String lastFieldName) {
    if (lastValue < firstValue) {
      throw new IllegalArgumentException(
          ExcelIndexDisplay.mustNotBeLessThan(
              lastFieldName, lastValue, firstFieldName, firstValue));
    }
  }
}
