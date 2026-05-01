package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.PaneType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** ExcelSheet pane, merge, sizing, and print-layout coverage. */
class ExcelSheetLayoutCoverageTest extends ExcelSheetTestSupport {
  @Test
  void autoSizeColumnsProducesDeterministicWidthsFromDisplayedContent() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.text("ID"));
      sheet.setCell("B1", ExcelCellValue.text("Riga onboarding checklist"));
      sheet.setCell("C1", ExcelCellValue.text("Q1"));
      sheet.autoSizeColumns();

      WorkbookSheetResult.SheetLayout layout = sheet.layout();
      assertTrue(
          layout.columns().get(1).widthCharacters() > layout.columns().get(0).widthCharacters());
      assertTrue(
          layout.columns().get(1).widthCharacters() > layout.columns().get(2).widthCharacters());
      assertTrue(layout.columns().get(1).widthCharacters() >= 20.0d);
    }
  }

  @Test
  void managesMergedRegionsSizingPaneZoomAndPrintLayout() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertSame(sheet, sheet.mergeCells("B2:A1"));
      assertEquals(1, poiSheet.getNumMergedRegions());
      assertEquals("A1:B2", poiSheet.getMergedRegion(0).formatAsString());
      assertSame(sheet, sheet.mergeCells("A1:B2"));
      assertEquals(1, poiSheet.getNumMergedRegions());
      assertThrows(IllegalArgumentException.class, () -> sheet.mergeCells("B2:C3"));

      assertSame(sheet, sheet.setColumnWidth(0, 1, 16.0));
      assertEquals(4096, poiSheet.getColumnWidth(0));
      assertEquals(4096, poiSheet.getColumnWidth(1));

      assertSame(sheet, sheet.setRowHeight(0, 1, 28.5));
      assertEquals((short) 570, poiSheet.getRow(0).getHeight());
      assertEquals((short) 570, poiSheet.getRow(1).getHeight());
      sheet.setRowHeight(0, 0, ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS);
      assertEquals(
          ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS,
          sheet.layout().rows().get(0).heightPoints());

      assertSame(sheet, sheet.setPane(new ExcelSheetPane.Frozen(1, 2, 3, 4)));
      assertNotNull(poiSheet.getPaneInformation());
      assertEquals(1, poiSheet.getPaneInformation().getVerticalSplitPosition());
      assertEquals(2, poiSheet.getPaneInformation().getHorizontalSplitPosition());
      assertEquals(3, poiSheet.getPaneInformation().getVerticalSplitLeftColumn());
      assertEquals(4, poiSheet.getPaneInformation().getHorizontalSplitTopRow());
      assertSame(sheet, sheet.setPane(new ExcelSheetPane.Frozen(0, 2, 0, 2)));
      assertEquals(0, poiSheet.getPaneInformation().getVerticalSplitPosition());
      assertEquals(2, poiSheet.getPaneInformation().getHorizontalSplitPosition());
      assertEquals(0, poiSheet.getPaneInformation().getVerticalSplitLeftColumn());
      assertEquals(2, poiSheet.getPaneInformation().getHorizontalSplitTopRow());
      assertSame(sheet, sheet.setPane(new ExcelSheetPane.Frozen(2, 0, 2, 0)));
      assertEquals(2, poiSheet.getPaneInformation().getVerticalSplitPosition());
      assertEquals(0, poiSheet.getPaneInformation().getHorizontalSplitPosition());
      assertEquals(2, poiSheet.getPaneInformation().getVerticalSplitLeftColumn());
      assertEquals(0, poiSheet.getPaneInformation().getHorizontalSplitTopRow());
      assertSame(
          sheet,
          sheet.setPane(new ExcelSheetPane.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT)));
      assertEquals(
          new ExcelSheetPane.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
          sheet.layout().pane());
      assertSame(sheet, sheet.setZoom(125));
      assertEquals(125, sheet.layout().zoomPercent());
      ExcelPrintLayout printLayout =
          new ExcelPrintLayout(
              new ExcelPrintLayout.Area.Range("A1:B20"),
              ExcelPrintOrientation.LANDSCAPE,
              new ExcelPrintLayout.Scaling.Fit(1, 0),
              new ExcelPrintLayout.TitleRows.Band(0, 0),
              new ExcelPrintLayout.TitleColumns.Band(0, 0),
              new ExcelHeaderFooterText("Budget", "", ""),
              new ExcelHeaderFooterText("", "Page &P", ""));
      assertSame(sheet, sheet.setPrintLayout(printLayout));
      assertEquals(printLayout, sheet.printLayout());
      assertSame(sheet, sheet.clearPrintLayout());
      ExcelPrintLayout clearedPrintLayout = sheet.printLayout();
      assertEquals(new ExcelPrintLayout.Area.None(), clearedPrintLayout.printArea());
      assertEquals(ExcelPrintOrientation.PORTRAIT, clearedPrintLayout.orientation());
      assertEquals(new ExcelPrintLayout.Scaling.Automatic(), clearedPrintLayout.scaling());
      assertEquals(new ExcelPrintLayout.TitleRows.None(), clearedPrintLayout.repeatingRows());
      assertEquals(new ExcelPrintLayout.TitleColumns.None(), clearedPrintLayout.repeatingColumns());
      assertEquals(ExcelHeaderFooterText.blank(), clearedPrintLayout.header());
      assertEquals(ExcelHeaderFooterText.blank(), clearedPrintLayout.footer());

      assertSame(sheet, sheet.unmergeCells("A1:B2"));
      assertEquals(0, poiSheet.getNumMergedRegions());
    }
  }

  @Test
  void validatesStructuralRangeAndSizingHelpers() {
    ExcelRange exactRange = ExcelRange.parse("A1:B2");
    CellRangeAddress exactAddress = new CellRangeAddress(0, 1, 0, 1);

    assertTrue(ExcelSheet.matches(exactAddress, exactRange));
    assertFalse(ExcelSheet.matches(exactAddress, ExcelRange.parse("A2:B3")));
    assertFalse(ExcelSheet.matches(exactAddress, ExcelRange.parse("A1:B3")));
    assertFalse(ExcelSheet.matches(exactAddress, ExcelRange.parse("B1:C2")));
    assertFalse(ExcelSheet.matches(exactAddress, ExcelRange.parse("A1:C2")));

    assertTrue(ExcelSheet.intersects(exactAddress, ExcelRange.parse("B2:C3")));
    assertFalse(ExcelSheet.intersects(new CellRangeAddress(3, 4, 0, 1), ExcelRange.parse("A1:B2")));
    assertFalse(ExcelSheet.intersects(new CellRangeAddress(0, 1, 0, 1), ExcelRange.parse("A4:B5")));
    assertFalse(ExcelSheet.intersects(new CellRangeAddress(0, 1, 3, 4), ExcelRange.parse("A1:B2")));
    assertFalse(ExcelSheet.intersects(new CellRangeAddress(0, 1, 0, 1), ExcelRange.parse("C1:D2")));

    assertEquals(4096, ExcelSheet.toColumnWidthUnits(16.0d));
    assertThrows(IllegalArgumentException.class, () -> ExcelSheet.toColumnWidthUnits(256.0d));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelSheet.toColumnWidthUnits(Double.MIN_VALUE));

    assertEquals(28.5f, ExcelSheet.toRowHeightPoints(28.5d));
    assertDoesNotThrow(
        () -> ExcelSheet.toRowHeightPoints(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelSheet.toRowHeightPoints(
                Math.nextUp(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS)));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelSheet.toRowHeightPoints(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS + 1.0d));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelSheet.toRowHeightPoints(Double.MIN_VALUE));
  }

  @Test
  void validatesPaneRegionAndPrintLayoutValueTypes() {
    assertDoesNotThrow(() -> new ExcelSheetPane.Frozen(1, 2, 1, 2));
    assertDoesNotThrow(() -> new ExcelSheetPane.Frozen(0, 2, 0, 2));
    assertDoesNotThrow(() -> new ExcelSheetPane.Frozen(2, 0, 2, 0));
    assertDoesNotThrow(() -> new ExcelSheetPane.Split(0, 1200, 0, 3, ExcelPaneRegion.LOWER_LEFT));
    assertDoesNotThrow(() -> new ExcelSheetPane.Split(1200, 0, 3, 0, ExcelPaneRegion.UPPER_RIGHT));
    assertDoesNotThrow(
        () ->
            new ExcelPrintLayout(
                new ExcelPrintLayout.Area.None(),
                ExcelPrintOrientation.PORTRAIT,
                new ExcelPrintLayout.Scaling.Automatic(),
                new ExcelPrintLayout.TitleRows.None(),
                new ExcelPrintLayout.TitleColumns.None(),
                ExcelHeaderFooterText.blank(),
                ExcelHeaderFooterText.blank()));
    assertDoesNotThrow(() -> ExcelSheetViewSupport.requireZoomPercent(10));
    assertDoesNotThrow(() -> ExcelSheetViewSupport.requireZoomPercent(400));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(0, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(0, 1, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(1, 0, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(2, 1, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(1, 2, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelSheetPane.Split(0, 1200, 1, 0, ExcelPaneRegion.LOWER_LEFT));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelSheetPane.Split(1200, 0, 0, 1, ExcelPaneRegion.LOWER_LEFT));
    assertThrows(NullPointerException.class, () -> new ExcelSheetPane.Split(1200, 0, 0, 0, null));
    assertTrue(ExcelHeaderFooterText.blank().isBlank());
    assertFalse(new ExcelHeaderFooterText("Left", "", "").isBlank());
    assertFalse(new ExcelHeaderFooterText("", "Center", "").isBlank());
    assertFalse(new ExcelHeaderFooterText("", "", "Right").isBlank());
    assertThrows(NullPointerException.class, () -> new ExcelHeaderFooterText(null, "", ""));
    assertThrows(IllegalArgumentException.class, () -> new ExcelPrintLayout.Area.Range(" "));
    assertThrows(IllegalArgumentException.class, () -> new ExcelPrintLayout.Scaling.Fit(-1, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelPrintLayout.Scaling.Fit(Short.MAX_VALUE + 1, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleRows.Band(-1, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleRows.Band(0, -1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleRows.Band(2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintLayout.TitleRows.Band(
                0, org.apache.poi.ss.SpreadsheetVersion.EXCEL2007.getLastRowIndex() + 1));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleColumns.Band(-1, 0));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleColumns.Band(0, -1));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleColumns.Band(2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintLayout.TitleColumns.Band(
                0, org.apache.poi.ss.SpreadsheetVersion.EXCEL2007.getLastColumnIndex() + 1));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPrintLayout(
                null,
                ExcelPrintOrientation.PORTRAIT,
                new ExcelPrintLayout.Scaling.Automatic(),
                new ExcelPrintLayout.TitleRows.None(),
                new ExcelPrintLayout.TitleColumns.None(),
                ExcelHeaderFooterText.blank(),
                ExcelHeaderFooterText.blank()));
    assertEquals(ExcelPaneRegion.UPPER_LEFT, ExcelPanePoiBridge.fromPoi(PaneType.UPPER_LEFT));
    assertEquals(ExcelPaneRegion.UPPER_RIGHT, ExcelPanePoiBridge.fromPoi(PaneType.UPPER_RIGHT));
    assertEquals(ExcelPaneRegion.LOWER_LEFT, ExcelPanePoiBridge.fromPoi(PaneType.LOWER_LEFT));
    assertEquals(ExcelPaneRegion.LOWER_RIGHT, ExcelPanePoiBridge.fromPoi(PaneType.LOWER_RIGHT));
    assertEquals(PaneType.UPPER_LEFT, ExcelPanePoiBridge.toPoi(ExcelPaneRegion.UPPER_LEFT));
    assertEquals(PaneType.UPPER_RIGHT, ExcelPanePoiBridge.toPoi(ExcelPaneRegion.UPPER_RIGHT));
    assertEquals(PaneType.LOWER_LEFT, ExcelPanePoiBridge.toPoi(ExcelPaneRegion.LOWER_LEFT));
    assertEquals(PaneType.LOWER_RIGHT, ExcelPanePoiBridge.toPoi(ExcelPaneRegion.LOWER_RIGHT));
    assertThrows(IllegalArgumentException.class, () -> ExcelSheetViewSupport.requireZoomPercent(9));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelSheetViewSupport.requireZoomPercent(401));
  }

  @Test
  void reportsExcelNativePrintTitleBandDiagnostics() {
    IllegalArgumentException negativeTitleRow =
        assertThrows(
            IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleRows.Band(-1, 0));
    assertTrue(negativeTitleRow.getMessage().contains("firstRowIndex -1"));
    assertTrue(negativeTitleRow.getMessage().contains("Excel row 1"));

    IllegalArgumentException negativeLastTitleRow =
        assertThrows(
            IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleRows.Band(0, -1));
    assertTrue(negativeLastTitleRow.getMessage().contains("lastRowIndex -1"));
    assertTrue(negativeLastTitleRow.getMessage().contains("Excel row 1"));

    IllegalArgumentException descendingTitleRow =
        assertThrows(
            IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleRows.Band(2, 1));
    assertTrue(descendingTitleRow.getMessage().contains("lastRowIndex 1 (Excel row 2)"));
    assertTrue(descendingTitleRow.getMessage().contains("firstRowIndex 2 (Excel row 3)"));

    IllegalArgumentException overflowTitleRow =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ExcelPrintLayout.TitleRows.Band(
                    0, org.apache.poi.ss.SpreadsheetVersion.EXCEL2007.getLastRowIndex() + 1));
    assertTrue(overflowTitleRow.getMessage().contains("lastRowIndex 1048576 (Excel row 1048577)"));
    assertTrue(overflowTitleRow.getMessage().contains("1048575 (Excel row 1048576)"));

    IllegalArgumentException negativeTitleColumn =
        assertThrows(
            IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleColumns.Band(-1, 0));
    assertTrue(negativeTitleColumn.getMessage().contains("firstColumnIndex -1"));
    assertTrue(negativeTitleColumn.getMessage().contains("Excel column A"));

    IllegalArgumentException negativeLastTitleColumn =
        assertThrows(
            IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleColumns.Band(0, -1));
    assertTrue(negativeLastTitleColumn.getMessage().contains("lastColumnIndex -1"));
    assertTrue(negativeLastTitleColumn.getMessage().contains("Excel column A"));

    IllegalArgumentException descendingTitleColumn =
        assertThrows(
            IllegalArgumentException.class, () -> new ExcelPrintLayout.TitleColumns.Band(2, 1));
    assertTrue(descendingTitleColumn.getMessage().contains("lastColumnIndex 1 (Excel column B)"));
    assertTrue(descendingTitleColumn.getMessage().contains("firstColumnIndex 2 (Excel column C)"));

    IllegalArgumentException overflowTitleColumn =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ExcelPrintLayout.TitleColumns.Band(
                    0, org.apache.poi.ss.SpreadsheetVersion.EXCEL2007.getLastColumnIndex() + 1));
    assertTrue(
        overflowTitleColumn.getMessage().contains("lastColumnIndex 16384 (Excel column XFE)"));
    assertTrue(overflowTitleColumn.getMessage().contains("16383 (Excel column XFD)"));
  }

  @Test
  void validatesMergedRegionLookupAndOverlapHelpers() throws Exception {
    ExcelRange exactRange = ExcelRange.parse("A1:B2");
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet emptySheet = poiWorkbook.createSheet("Empty");
      Sheet mergedSheet = poiWorkbook.createSheet("Merged");
      mergedSheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 1));
      mergedSheet.addMergedRegion(new CellRangeAddress(3, 4, 3, 4));

      assertEquals(-1, ExcelSheet.findMergedRegionIndex(emptySheet, exactRange));
      assertEquals(0, ExcelSheet.findMergedRegionIndex(mergedSheet, exactRange));
      assertEquals(1, ExcelSheet.findMergedRegionIndex(mergedSheet, ExcelRange.parse("D4:E5")));
      assertEquals(-1, ExcelSheet.findMergedRegionIndex(mergedSheet, ExcelRange.parse("A1:B3")));

      assertDoesNotThrow(
          () -> ExcelSheet.requireNoMergedRegionOverlap(mergedSheet, ExcelRange.parse("G1:H2")));
      IllegalArgumentException overlap =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelSheet.requireNoMergedRegionOverlap(mergedSheet, ExcelRange.parse("B2:C3")));
      assertEquals("Merged range overlaps existing merged region: A1:B2", overlap.getMessage());
    }
  }
}
