package dev.erst.gridgrind.engine.backend.xssf;

import static dev.erst.gridgrind.engine.backend.xssf.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.vocabulary.ExcelPaneRegion;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.engine.layout.ColumnAdapter;
import dev.erst.gridgrind.engine.layout.LayoutAdapter;
import dev.erst.gridgrind.engine.layout.RowAdapter;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import dev.erst.gridgrind.engine.styles.StyleAdapter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.PaneType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** XssfSheet pane, merge, sizing, and print-layout coverage. */
class ExcelSheetLayoutCoverageTest extends ExcelSheetTestSupport {
  @Test
  void autoSizeColumnsProducesDeterministicWidthsFromDisplayedContent() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      XssfSheet sheet = new XssfSheet(poiSheet, new StyleAdapter(poiWorkbook), evaluator);

      sheet.cells().setCell("A1", new CellValue.Text("ID"));
      sheet.cells().setCell("B1", new CellValue.Text("Riga onboarding checklist"));
      sheet.cells().setCell("C1", new CellValue.Text("Q1"));
      sheet.columns().autoSize();

      SheetLayoutReport layout = sheet.layout().snapshot();
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
      XssfSheet sheet = new XssfSheet(poiSheet, new StyleAdapter(poiWorkbook), evaluator);
      LayoutAdapter layout = sheet.layout();
      ColumnAdapter columns = sheet.columns();
      RowAdapter rows = sheet.rows();

      assertSame(layout, layout.mergeCells("B2:A1"));
      assertEquals(1, poiSheet.getNumMergedRegions());
      assertEquals("A1:B2", poiSheet.getMergedRegion(0).formatAsString());
      assertSame(layout, layout.mergeCells("A1:B2"));
      assertEquals(1, poiSheet.getNumMergedRegions());
      assertThrows(IllegalArgumentException.class, () -> layout.mergeCells("B2:C3"));

      assertSame(columns, columns.setWidth(0, 1, 16.0));
      assertEquals(4096, poiSheet.getColumnWidth(0));
      assertEquals(4096, poiSheet.getColumnWidth(1));

      assertSame(rows, rows.setHeight(0, 1, 28.5));
      assertEquals((short) 570, poiSheet.getRow(0).getHeight());
      assertEquals((short) 570, poiSheet.getRow(1).getHeight());
      rows.setHeight(0, 0, ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS);
      assertEquals(
          ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS,
          layout.snapshot().rows().get(0).heightPoints());

      assertSame(layout, layout.setPane(new PaneSpec.Frozen(1, 2, 3, 4)));
      assertNotNull(poiSheet.getPaneInformation());
      assertEquals(1, poiSheet.getPaneInformation().getVerticalSplitPosition());
      assertEquals(2, poiSheet.getPaneInformation().getHorizontalSplitPosition());
      assertEquals(3, poiSheet.getPaneInformation().getVerticalSplitLeftColumn());
      assertEquals(4, poiSheet.getPaneInformation().getHorizontalSplitTopRow());
      assertSame(layout, layout.setPane(new PaneSpec.Frozen(0, 2, 0, 2)));
      assertEquals(0, poiSheet.getPaneInformation().getVerticalSplitPosition());
      assertEquals(2, poiSheet.getPaneInformation().getHorizontalSplitPosition());
      assertEquals(0, poiSheet.getPaneInformation().getVerticalSplitLeftColumn());
      assertEquals(2, poiSheet.getPaneInformation().getHorizontalSplitTopRow());
      assertSame(layout, layout.setPane(new PaneSpec.Frozen(2, 0, 2, 0)));
      assertEquals(2, poiSheet.getPaneInformation().getVerticalSplitPosition());
      assertEquals(0, poiSheet.getPaneInformation().getHorizontalSplitPosition());
      assertEquals(2, poiSheet.getPaneInformation().getVerticalSplitLeftColumn());
      assertEquals(0, poiSheet.getPaneInformation().getHorizontalSplitTopRow());
      assertSame(
          layout,
          layout.setPane(new PaneSpec.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT)));
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.PaneReport.Split(
              1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
          layout.snapshot().pane());
      assertSame(layout, layout.setZoom(125));
      assertEquals(125, layout.snapshot().zoomPercent());
      ResolvedPrintLayout printLayout =
          new ResolvedPrintLayout(
              new PrintAreaSpec.Range(new RangeRef("A1:B20")),
              ExcelPrintOrientation.LANDSCAPE,
              new PrintScalingSpec.Fit(1, 0),
              new PrintTitleRowsSpec.Band(0, 0),
              new PrintTitleColumnsSpec.Band(0, 0),
              new HeaderFooterTextReport("Budget", "", ""),
              new HeaderFooterTextReport("", "Page &P", ""));
      assertSame(layout, layout.setPrintLayout(printLayout));
      assertEquals(printLayout, layout.printLayout());
      assertSame(layout, layout.clearPrintLayout());
      ResolvedPrintLayout clearedPrintLayout = layout.printLayout();
      assertEquals(new PrintAreaSpec.None(), clearedPrintLayout.printArea());
      assertEquals(ExcelPrintOrientation.PORTRAIT, clearedPrintLayout.orientation());
      assertEquals(new PrintScalingSpec.Automatic(), clearedPrintLayout.scaling());
      assertEquals(new PrintTitleRowsSpec.None(), clearedPrintLayout.repeatingRows());
      assertEquals(new PrintTitleColumnsSpec.None(), clearedPrintLayout.repeatingColumns());
      assertEquals(HeaderFooterTextReport.blank(), clearedPrintLayout.header());
      assertEquals(HeaderFooterTextReport.blank(), clearedPrintLayout.footer());

      assertSame(layout, layout.unmergeCells("A1:B2"));
      assertEquals(0, poiSheet.getNumMergedRegions());
    }
  }

  @Test
  void validatesStructuralRangeAndSizingHelpers() {
    CellBounds exactRange = CellBounds.parse("A1:B2");
    CellRangeAddress exactAddress = new CellRangeAddress(0, 1, 0, 1);

    assertTrue(SheetStructure.matches(exactAddress, exactRange));
    assertFalse(SheetStructure.matches(exactAddress, CellBounds.parse("A2:B3")));
    assertFalse(SheetStructure.matches(exactAddress, CellBounds.parse("A1:B3")));
    assertFalse(SheetStructure.matches(exactAddress, CellBounds.parse("B1:C2")));
    assertFalse(SheetStructure.matches(exactAddress, CellBounds.parse("A1:C2")));

    assertTrue(SheetStructure.intersects(exactAddress, CellBounds.parse("B2:C3")));
    assertFalse(
        SheetStructure.intersects(new CellRangeAddress(3, 4, 0, 1), CellBounds.parse("A1:B2")));
    assertFalse(
        SheetStructure.intersects(new CellRangeAddress(0, 1, 0, 1), CellBounds.parse("A4:B5")));
    assertFalse(
        SheetStructure.intersects(new CellRangeAddress(0, 1, 3, 4), CellBounds.parse("A1:B2")));
    assertFalse(
        SheetStructure.intersects(new CellRangeAddress(0, 1, 0, 1), CellBounds.parse("C1:D2")));

    assertEquals(4096, SheetStructure.toColumnWidthUnits(16.0d));
    assertThrows(IllegalArgumentException.class, () -> SheetStructure.toColumnWidthUnits(256.0d));
    assertThrows(
        IllegalArgumentException.class, () -> SheetStructure.toColumnWidthUnits(Double.MIN_VALUE));

    assertEquals(28.5f, SheetStructure.toRowHeightPoints(28.5d));
    assertDoesNotThrow(
        () -> SheetStructure.toRowHeightPoints(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            SheetStructure.toRowHeightPoints(
                Math.nextUp(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            SheetStructure.toRowHeightPoints(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS + 1.0d));
    assertThrows(
        IllegalArgumentException.class, () -> SheetStructure.toRowHeightPoints(Double.MIN_VALUE));
  }

  @Test
  void validatesPaneRegionAndPrintLayoutValueTypes() {
    assertDoesNotThrow(() -> new PaneSpec.Frozen(1, 2, 1, 2));
    assertDoesNotThrow(() -> new PaneSpec.Frozen(0, 2, 0, 2));
    assertDoesNotThrow(() -> new PaneSpec.Frozen(2, 0, 2, 0));
    assertDoesNotThrow(() -> new PaneSpec.Split(0, 1200, 0, 3, ExcelPaneRegion.LOWER_LEFT));
    assertDoesNotThrow(() -> new PaneSpec.Split(1200, 0, 3, 0, ExcelPaneRegion.UPPER_RIGHT));
    assertDoesNotThrow(
        () ->
            new ResolvedPrintLayout(
                new PrintAreaSpec.None(),
                ExcelPrintOrientation.PORTRAIT,
                new PrintScalingSpec.Automatic(),
                new PrintTitleRowsSpec.None(),
                new PrintTitleColumnsSpec.None(),
                HeaderFooterTextReport.blank(),
                HeaderFooterTextReport.blank()));
    assertDoesNotThrow(() -> LayoutAdapter.requireZoomPercent(10));
    assertDoesNotThrow(() -> LayoutAdapter.requireZoomPercent(400));
    assertThrows(IllegalArgumentException.class, () -> new PaneSpec.Frozen(0, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new PaneSpec.Frozen(0, 1, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new PaneSpec.Frozen(1, 0, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new PaneSpec.Frozen(2, 1, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new PaneSpec.Frozen(1, 2, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PaneSpec.Split(0, 1200, 1, 0, ExcelPaneRegion.LOWER_LEFT));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PaneSpec.Split(1200, 0, 0, 1, ExcelPaneRegion.LOWER_LEFT));
    assertThrows(NullPointerException.class, () -> new PaneSpec.Split(1200, 0, 0, 0, null));
    assertTrue(HeaderFooterTextReport.blank().empty());
    assertFalse(new HeaderFooterTextReport("Left", "", "").empty());
    assertFalse(new HeaderFooterTextReport("", "Center", "").empty());
    assertFalse(new HeaderFooterTextReport("", "", "Right").empty());
    assertThrows(NullPointerException.class, () -> new HeaderFooterTextReport(null, "", ""));
    assertThrows(IllegalArgumentException.class, () -> new PrintAreaSpec.Range(new RangeRef(" ")));
    assertThrows(IllegalArgumentException.class, () -> new PrintScalingSpec.Fit(-1, 0));
    assertThrows(
        IllegalArgumentException.class, () -> new PrintScalingSpec.Fit(Short.MAX_VALUE + 1, 0));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsSpec.Band(-1, 0));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsSpec.Band(0, -1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsSpec.Band(2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintTitleRowsSpec.Band(
                0, org.apache.poi.ss.SpreadsheetVersion.EXCEL2007.getLastRowIndex() + 1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsSpec.Band(-1, 0));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsSpec.Band(0, -1));
    assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsSpec.Band(2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintTitleColumnsSpec.Band(
                0, org.apache.poi.ss.SpreadsheetVersion.EXCEL2007.getLastColumnIndex() + 1));
    assertThrows(
        NullPointerException.class,
        () ->
            new ResolvedPrintLayout(
                null,
                ExcelPrintOrientation.PORTRAIT,
                new PrintScalingSpec.Automatic(),
                new PrintTitleRowsSpec.None(),
                new PrintTitleColumnsSpec.None(),
                HeaderFooterTextReport.blank(),
                HeaderFooterTextReport.blank()));
    assertEquals(ExcelPaneRegion.UPPER_LEFT, LayoutAdapter.fromPoiPane(PaneType.UPPER_LEFT));
    assertEquals(ExcelPaneRegion.UPPER_RIGHT, LayoutAdapter.fromPoiPane(PaneType.UPPER_RIGHT));
    assertEquals(ExcelPaneRegion.LOWER_LEFT, LayoutAdapter.fromPoiPane(PaneType.LOWER_LEFT));
    assertEquals(ExcelPaneRegion.LOWER_RIGHT, LayoutAdapter.fromPoiPane(PaneType.LOWER_RIGHT));
    assertEquals(PaneType.UPPER_LEFT, LayoutAdapter.toPoiPane(ExcelPaneRegion.UPPER_LEFT));
    assertEquals(PaneType.UPPER_RIGHT, LayoutAdapter.toPoiPane(ExcelPaneRegion.UPPER_RIGHT));
    assertEquals(PaneType.LOWER_LEFT, LayoutAdapter.toPoiPane(ExcelPaneRegion.LOWER_LEFT));
    assertEquals(PaneType.LOWER_RIGHT, LayoutAdapter.toPoiPane(ExcelPaneRegion.LOWER_RIGHT));
    assertThrows(IllegalArgumentException.class, () -> LayoutAdapter.requireZoomPercent(9));
    assertThrows(IllegalArgumentException.class, () -> LayoutAdapter.requireZoomPercent(401));
  }

  @Test
  void reportsExcelNativePrintTitleBandDiagnostics() {
    IllegalArgumentException negativeTitleRow =
        assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsSpec.Band(-1, 0));
    assertTrue(negativeTitleRow.getMessage().contains("firstRowIndex -1"));
    assertTrue(negativeTitleRow.getMessage().contains("Excel row 1"));

    IllegalArgumentException negativeLastTitleRow =
        assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsSpec.Band(0, -1));
    assertTrue(negativeLastTitleRow.getMessage().contains("lastRowIndex -1"));
    assertTrue(negativeLastTitleRow.getMessage().contains("Excel row 1"));

    IllegalArgumentException descendingTitleRow =
        assertThrows(IllegalArgumentException.class, () -> new PrintTitleRowsSpec.Band(2, 1));
    assertTrue(descendingTitleRow.getMessage().contains("lastRowIndex 1 (Excel row 2)"));
    assertTrue(descendingTitleRow.getMessage().contains("firstRowIndex 2 (Excel row 3)"));

    IllegalArgumentException overflowTitleRow =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new PrintTitleRowsSpec.Band(
                    0, org.apache.poi.ss.SpreadsheetVersion.EXCEL2007.getLastRowIndex() + 1));
    assertTrue(overflowTitleRow.getMessage().contains("lastRowIndex 1048576 (Excel row 1048577)"));
    assertTrue(overflowTitleRow.getMessage().contains("1048575 (Excel row 1048576)"));

    IllegalArgumentException negativeTitleColumn =
        assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsSpec.Band(-1, 0));
    assertTrue(negativeTitleColumn.getMessage().contains("firstColumnIndex -1"));
    assertTrue(negativeTitleColumn.getMessage().contains("Excel column A"));

    IllegalArgumentException negativeLastTitleColumn =
        assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsSpec.Band(0, -1));
    assertTrue(negativeLastTitleColumn.getMessage().contains("lastColumnIndex -1"));
    assertTrue(negativeLastTitleColumn.getMessage().contains("Excel column A"));

    IllegalArgumentException descendingTitleColumn =
        assertThrows(IllegalArgumentException.class, () -> new PrintTitleColumnsSpec.Band(2, 1));
    assertTrue(descendingTitleColumn.getMessage().contains("lastColumnIndex 1 (Excel column B)"));
    assertTrue(descendingTitleColumn.getMessage().contains("firstColumnIndex 2 (Excel column C)"));

    IllegalArgumentException overflowTitleColumn =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new PrintTitleColumnsSpec.Band(
                    0, org.apache.poi.ss.SpreadsheetVersion.EXCEL2007.getLastColumnIndex() + 1));
    assertTrue(
        overflowTitleColumn.getMessage().contains("lastColumnIndex 16384 (Excel column XFE)"));
    assertTrue(overflowTitleColumn.getMessage().contains("16383 (Excel column XFD)"));
  }

  @Test
  void validatesMergedRegionLookupAndOverlapHelpers() throws Exception {
    CellBounds exactRange = CellBounds.parse("A1:B2");
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet emptySheet = poiWorkbook.createSheet("Empty");
      Sheet mergedSheet = poiWorkbook.createSheet("Merged");
      mergedSheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 1));
      mergedSheet.addMergedRegion(new CellRangeAddress(3, 4, 3, 4));

      assertEquals(-1, SheetStructure.findMergedRegionIndex(emptySheet, exactRange));
      assertEquals(0, SheetStructure.findMergedRegionIndex(mergedSheet, exactRange));
      assertEquals(1, SheetStructure.findMergedRegionIndex(mergedSheet, CellBounds.parse("D4:E5")));
      assertEquals(
          -1, SheetStructure.findMergedRegionIndex(mergedSheet, CellBounds.parse("A1:B3")));

      assertDoesNotThrow(
          () ->
              SheetStructure.requireNoMergedRegionOverlap(mergedSheet, CellBounds.parse("G1:H2")));
      IllegalArgumentException overlap =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  SheetStructure.requireNoMergedRegionOverlap(
                      mergedSheet, CellBounds.parse("B2:C3")));
      assertEquals("Merged range overlaps existing merged region: A1:B2", overlap.getMessage());
    }
  }
}
