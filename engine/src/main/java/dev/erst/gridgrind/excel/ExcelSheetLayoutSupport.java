package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Layout, merge, view, and print operations for one sheet wrapper. */
final class ExcelSheetLayoutSupport {
  private final Sheet sheet;
  private final ExcelPrintLayoutController printLayoutController;
  private final ExcelSheetPresentationController sheetPresentationController;
  private final ExcelRowStructureController rowStructureController;
  private final ExcelColumnStructureController columnStructureController;

  ExcelSheetLayoutSupport(
      Sheet sheet,
      ExcelPrintLayoutController printLayoutController,
      ExcelSheetPresentationController sheetPresentationController,
      ExcelRowStructureController rowStructureController,
      ExcelColumnStructureController columnStructureController) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.printLayoutController =
        Objects.requireNonNull(printLayoutController, "printLayoutController must not be null");
    this.sheetPresentationController =
        Objects.requireNonNull(
            sheetPresentationController, "sheetPresentationController must not be null");
    this.rowStructureController =
        Objects.requireNonNull(rowStructureController, "rowStructureController must not be null");
    this.columnStructureController =
        Objects.requireNonNull(
            columnStructureController, "columnStructureController must not be null");
  }

  void mergeCells(String range) {
    ExcelArgumentSupport.requireNonBlank(range, "range");
    ExcelRange excelRange = ExcelRange.parse(range);
    requireMergeableRange(range, excelRange);
    if (ExcelSheetStructureSupport.findMergedRegionIndex(sheet, excelRange) >= 0) {
      return;
    }
    ExcelSheetStructureSupport.requireNoMergedRegionOverlap(sheet, excelRange);
    sheet.addMergedRegion(ExcelSheetStructureSupport.toCellRangeAddress(excelRange));
  }

  void unmergeCells(String range) {
    ExcelArgumentSupport.requireNonBlank(range, "range");
    ExcelRange excelRange = ExcelRange.parse(range);
    int mergedRegionIndex = ExcelSheetStructureSupport.findMergedRegionIndex(sheet, excelRange);
    if (mergedRegionIndex < 0) {
      throw new IllegalArgumentException("No merged region matches range: " + range);
    }
    sheet.removeMergedRegion(mergedRegionIndex);
  }

  void setPane(ExcelSheetPane pane) {
    Objects.requireNonNull(pane, "pane must not be null");
    ExcelSheetViewSupport.setPane(xssfSheet(), pane);
  }

  void setZoom(int zoomPercent) {
    ExcelSheetViewSupport.requireZoomPercent(zoomPercent);
    ExcelSheetViewSupport.setZoomPercent(xssfSheet(), zoomPercent);
  }

  void setPresentation(ExcelSheetPresentation presentation) {
    Objects.requireNonNull(presentation, "presentation must not be null");
    sheetPresentationController.setPresentation(xssfSheet(), presentation);
  }

  void setPrintLayout(ExcelPrintLayout printLayout) {
    Objects.requireNonNull(printLayout, "printLayout must not be null");
    printLayoutController.setPrintLayout(xssfSheet(), printLayout);
  }

  void clearPrintLayout() {
    printLayoutController.clearPrintLayout(xssfSheet());
  }

  List<WorkbookSheetResult.MergedRegion> mergedRegions() {
    List<WorkbookSheetResult.MergedRegion> mergedRegions =
        new ArrayList<>(sheet.getNumMergedRegions());
    for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
      mergedRegions.add(
          new WorkbookSheetResult.MergedRegion(
              sheet.getMergedRegion(regionIndex).formatAsString()));
    }
    return List.copyOf(mergedRegions);
  }

  WorkbookSheetResult.SheetLayout layout(String sheetName) {
    return new WorkbookSheetResult.SheetLayout(
        sheetName,
        ExcelSheetViewSupport.pane(xssfSheet()),
        ExcelSheetViewSupport.zoomPercent(xssfSheet()),
        sheetPresentationController.presentation(xssfSheet()),
        columnStructureController.columnLayouts(xssfSheet()),
        rowStructureController.rowLayouts(xssfSheet()));
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

  private static void requireMergeableRange(String range, ExcelRange excelRange) {
    if (excelRange.rowCount() == 1 && excelRange.columnCount() == 1) {
      throw new IllegalArgumentException("range must span at least two cells: " + range);
    }
  }
}
