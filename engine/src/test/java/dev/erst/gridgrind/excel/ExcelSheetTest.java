package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Proxy;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Integration tests for ExcelSheet typed reads, writes, and previews. */
class ExcelSheetTest {
  @Test
  void readsWritesAndSnapshotsTypedCellsAndFormulaResults() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertSame(
          sheet,
          sheet.appendRow(
              ExcelCellValue.text("Name"),
              ExcelCellValue.number(42.5),
              ExcelCellValue.bool(true),
              ExcelCellValue.formula("B1*2"),
              ExcelCellValue.formula("TRUE()"),
              ExcelCellValue.formula("\"Hi\"")));
      sheet.appendRow(
          ExcelCellValue.blank(),
          ExcelCellValue.date(LocalDate.of(2026, 3, 23)),
          ExcelCellValue.dateTime(LocalDateTime.of(2026, 3, 23, 14, 15, 16)),
          ExcelCellValue.formula("1/0"));
      sheet.setCell("B3", ExcelCellValue.date(LocalDate.of(2026, 3, 24)));
      sheet.setCell("C3", ExcelCellValue.dateTime(LocalDateTime.of(2026, 3, 24, 9, 0)));
      sheet.autoSizeColumns();

      Row errorRow = poiSheet.createRow(3);
      Cell errorCell = errorRow.createCell(0);
      errorCell.setCellErrorValue(FormulaError.DIV0.getCode());

      assertEquals("Budget", sheet.name());
      assertEquals("Name", sheet.text("A1"));
      assertEquals(85.0, sheet.number("D1"));
      assertTrue(sheet.bool("C1"));
      assertTrue(sheet.bool("E1"));
      assertEquals("B1*2", sheet.formula("D1"));
      assertEquals(4, sheet.physicalRowCount());
      assertEquals(3, sheet.lastRowIndex());
      assertEquals(5, sheet.lastColumnIndex());

      ExcelCellSnapshot.TextSnapshot textSnapshot =
          (ExcelCellSnapshot.TextSnapshot) sheet.snapshotCell("A1");
      assertEquals("STRING", textSnapshot.declaredType());
      assertEquals("STRING", textSnapshot.effectiveType());
      assertEquals("Name", textSnapshot.stringValue());

      ExcelCellSnapshot.NumberSnapshot numberSnapshot =
          (ExcelCellSnapshot.NumberSnapshot) sheet.snapshotCell("B1");
      assertEquals("NUMERIC", numberSnapshot.declaredType());
      assertEquals("NUMERIC", numberSnapshot.effectiveType());
      assertEquals(42.5, numberSnapshot.numberValue());

      ExcelCellSnapshot.BooleanSnapshot booleanSnapshot =
          (ExcelCellSnapshot.BooleanSnapshot) sheet.snapshotCell("C1");
      assertEquals("BOOLEAN", booleanSnapshot.declaredType());
      assertEquals("BOOLEAN", booleanSnapshot.effectiveType());
      assertTrue(booleanSnapshot.booleanValue());

      ExcelCellSnapshot.BlankSnapshot blankSnapshot =
          (ExcelCellSnapshot.BlankSnapshot) sheet.snapshotCell("A2");
      assertEquals("BLANK", blankSnapshot.declaredType());
      assertEquals("BLANK", blankSnapshot.effectiveType());

      ExcelCellSnapshot.FormulaSnapshot stringFormulaSnapshot =
          (ExcelCellSnapshot.FormulaSnapshot) sheet.snapshotCell("F1");
      assertEquals("FORMULA", stringFormulaSnapshot.declaredType());
      assertEquals("FORMULA", stringFormulaSnapshot.effectiveType());
      assertEquals("\"Hi\"", stringFormulaSnapshot.formula());
      assertEquals(
          "Hi",
          ((ExcelCellSnapshot.TextSnapshot) stringFormulaSnapshot.evaluation()).stringValue());

      ExcelCellSnapshot.FormulaSnapshot errorFormulaSnapshot =
          (ExcelCellSnapshot.FormulaSnapshot) sheet.snapshotCell("D2");
      assertEquals("FORMULA", errorFormulaSnapshot.declaredType());
      assertEquals("FORMULA", errorFormulaSnapshot.effectiveType());
      assertEquals(
          "#DIV/0!",
          ((ExcelCellSnapshot.ErrorSnapshot) errorFormulaSnapshot.evaluation()).errorValue());

      ExcelCellSnapshot.ErrorSnapshot errorSnapshot =
          (ExcelCellSnapshot.ErrorSnapshot) sheet.snapshotCell("A4");
      assertEquals("ERROR", errorSnapshot.declaredType());
      assertEquals("ERROR", errorSnapshot.effectiveType());
      assertEquals("#DIV/0!", errorSnapshot.errorValue());

      List<ExcelPreviewRow> preview = sheet.preview(4, 6);
      assertEquals(4, preview.size());
      assertEquals("A1", preview.get(0).cells().get(0).address());
      assertTrue(preview.get(1).cells().stream().noneMatch(cell -> "A2".equals(cell.address())));
      assertEquals("Hi", preview.get(0).cells().get(5).displayValue());
    }
  }

  @Test
  void validatesWriteOperationArguments() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertThrows(NullPointerException.class, () -> sheet.setCell(null, ExcelCellValue.text("x")));
      assertThrows(
          IllegalArgumentException.class, () -> sheet.setCell(" ", ExcelCellValue.text("x")));
      assertThrows(NullPointerException.class, () -> sheet.setCell("A1", null));
      assertThrows(
          NullPointerException.class,
          () -> sheet.setRange(null, List.of(List.of(ExcelCellValue.text("x")))));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setRange(" ", List.of(List.of(ExcelCellValue.text("x")))));
      assertThrows(NullPointerException.class, () -> sheet.setRange("A1", null));
      assertThrows(IllegalArgumentException.class, () -> sheet.setRange("A1", List.of()));
      assertThrows(NullPointerException.class, () -> sheet.clearRange(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.clearRange(" "));
      assertThrows(NullPointerException.class, () -> sheet.mergeCells(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.mergeCells(" "));
      assertThrows(NullPointerException.class, () -> sheet.unmergeCells(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.unmergeCells(" "));
      assertThrows(IllegalArgumentException.class, () -> sheet.setColumnWidth(-1, 0, 16.0));
      assertThrows(IllegalArgumentException.class, () -> sheet.setColumnWidth(1, 0, 16.0));
      assertThrows(IllegalArgumentException.class, () -> sheet.setColumnWidth(0, 0, 0.0));
      assertThrows(IllegalArgumentException.class, () -> sheet.setColumnWidth(0, 0, 256.0));
      assertThrows(
          IllegalArgumentException.class, () -> sheet.setColumnWidth(0, 0, Double.MIN_VALUE));
      assertThrows(IllegalArgumentException.class, () -> sheet.setColumnWidth(0, 0, Double.NaN));
      assertThrows(IllegalArgumentException.class, () -> sheet.setRowHeight(-1, 0, 28.5));
      assertThrows(IllegalArgumentException.class, () -> sheet.setRowHeight(1, 0, 28.5));
      assertThrows(IllegalArgumentException.class, () -> sheet.setRowHeight(0, 0, 0.0));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setRowHeight(0, 0, (Short.MAX_VALUE / 20.0d) + 1.0d));
      assertThrows(
          IllegalArgumentException.class, () -> sheet.setRowHeight(0, 0, Double.MIN_VALUE));
      assertThrows(IllegalArgumentException.class, () -> sheet.setRowHeight(0, 0, Double.NaN));
      assertThrows(IllegalArgumentException.class, () -> sheet.freezePanes(-1, 0, 0, 0));
      assertThrows(IllegalArgumentException.class, () -> sheet.freezePanes(0, 0, 0, 0));
      assertThrows(IllegalArgumentException.class, () -> sheet.freezePanes(0, 1, 1, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.freezePanes(1, 0, 1, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.freezePanes(2, 1, 1, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.freezePanes(1, 2, 1, 1));
      assertThrows(
          NullPointerException.class,
          () -> sheet.applyStyle(null, ExcelCellStyle.numberFormat("0")));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.applyStyle(" ", ExcelCellStyle.numberFormat("0")));
      assertThrows(NullPointerException.class, () -> sheet.applyStyle("A1", null));
      assertThrows(NullPointerException.class, () -> sheet.appendRow((ExcelCellValue[]) null));
      assertThrows(
          NullPointerException.class, () -> sheet.appendRow(ExcelCellValue.text("x"), null));
    }
  }

  @Test
  void validatesAddressAndRangeParsing() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertThrows(NullPointerException.class, () -> sheet.snapshotCell(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.snapshotCell(" "));
      InvalidCellAddressException invalidSetCell =
          assertThrows(
              InvalidCellAddressException.class,
              () -> sheet.setCell(":", ExcelCellValue.text("x")));
      assertEquals(":", invalidSetCell.address());
      InvalidCellAddressException invalidSnapshotCell =
          assertThrows(InvalidCellAddressException.class, () -> sheet.snapshotCell(":"));
      assertEquals(":", invalidSnapshotCell.address());
      InvalidRangeAddressException invalidRangeSet =
          assertThrows(
              InvalidRangeAddressException.class,
              () -> sheet.setRange("A1:", List.of(List.of(ExcelCellValue.text("x")))));
      assertEquals("A1:", invalidRangeSet.range());
      InvalidRangeAddressException invalidRangeClear =
          assertThrows(InvalidRangeAddressException.class, () -> sheet.clearRange("A1:B2:C3"));
      assertEquals("A1:B2:C3", invalidRangeClear.range());
      InvalidRangeAddressException invalidRangeMerge =
          assertThrows(InvalidRangeAddressException.class, () -> sheet.mergeCells("A1:"));
      assertEquals("A1:", invalidRangeMerge.range());
      InvalidRangeAddressException invalidRangeUnmerge =
          assertThrows(InvalidRangeAddressException.class, () -> sheet.unmergeCells("A1:"));
      assertEquals("A1:", invalidRangeUnmerge.range());
      assertThrows(
          InvalidRangeAddressException.class,
          () -> sheet.applyStyle("A1:", ExcelCellStyle.numberFormat("0")));
      assertThrows(IllegalArgumentException.class, () -> sheet.mergeCells("A1"));
      assertThrows(IllegalArgumentException.class, () -> sheet.unmergeCells("A1:B2"));
    }
  }

  @Test
  void validatesPreviewAndReadFailures() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertThrows(IllegalArgumentException.class, () -> sheet.preview(0, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.preview(1, 0));
      assertEquals(List.of(), sheet.preview(3, 3));
      CellNotFoundException missingCell =
          assertThrows(CellNotFoundException.class, () -> sheet.text("A1"));
      assertEquals("A1", missingCell.address());
      assertThrows(IllegalArgumentException.class, () -> sheet.text(" "));

      sheet.setCell("A1", ExcelCellValue.text("Name"));
      sheet.setCell("B1", ExcelCellValue.formula("TRUE()"));
      sheet.setCell("C1", ExcelCellValue.formula("1+1"));

      assertThrows(CellNotFoundException.class, () -> sheet.text("B2"));
      assertThrows(CellNotFoundException.class, () -> sheet.text("D1"));
      assertThrows(IllegalStateException.class, () -> sheet.number("A1"));
      assertThrows(IllegalStateException.class, () -> sheet.number("B1"));
      assertThrows(IllegalStateException.class, () -> sheet.bool("C1"));
      assertThrows(IllegalStateException.class, () -> sheet.formula("A1"));
    }
  }

  @Test
  void supportsRangeWritesClearsStylesAndStyleAwarePreview() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertSame(
          sheet,
          sheet.setRange(
              "B2:A1",
              List.of(
                  List.of(ExcelCellValue.text("Item"), ExcelCellValue.number(42.0)),
                  List.of(ExcelCellValue.text("Tax"), ExcelCellValue.number(8.0)))));
      sheet.applyStyle(
          "A1:B1",
          new ExcelCellStyle(
              "#,##0.00",
              true,
              null,
              true,
              ExcelHorizontalAlignment.CENTER,
              ExcelVerticalAlignment.TOP));
      sheet.applyStyle("C1", ExcelCellStyle.emphasis(null, true));

      ExcelCellSnapshot styledValue = sheet.snapshotCell("A1");
      assertEquals("#,##0.00", styledValue.style().numberFormat());
      assertTrue(styledValue.style().bold());
      assertTrue(styledValue.style().wrapText());
      assertEquals(ExcelHorizontalAlignment.CENTER, styledValue.style().horizontalAlignment());
      assertEquals(ExcelVerticalAlignment.TOP, styledValue.style().verticalAlignment());

      List<ExcelPreviewRow> preview = sheet.preview(2, 3);
      assertTrue(preview.getFirst().cells().stream().anyMatch(cell -> "C1".equals(cell.address())));
      assertEquals("BLANK", sheet.snapshotCell("C1").effectiveType());
      assertTrue(sheet.snapshotCell("C1").style().italic());

      sheet.clearRange("A2:B2");

      ExcelCellSnapshot cleared = sheet.snapshotCell("A2");
      assertEquals("BLANK", cleared.declaredType());
      assertEquals("General", cleared.style().numberFormat());
      assertFalse(cleared.style().bold());
      assertEquals(ExcelHorizontalAlignment.GENERAL, cleared.style().horizontalAlignment());
      assertEquals(ExcelVerticalAlignment.BOTTOM, cleared.style().verticalAlignment());

      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setRange("A1:B2", List.of(List.of(ExcelCellValue.text("x")))));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setRange("A1:B2", List.of(List.of(), List.of(ExcelCellValue.text("x")))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet.setRange(
                  "A1:B2",
                  List.of(
                      List.of(ExcelCellValue.text("x")),
                      List.of(ExcelCellValue.text("y"), ExcelCellValue.text("z")))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet.setRange(
                  "A1:B2",
                  List.of(List.of(ExcelCellValue.text("x")), List.of(ExcelCellValue.text("y")))));
      List<List<ExcelCellValue>> rowsWithNull = new ArrayList<>();
      List<ExcelCellValue> rowWithNull = new ArrayList<>();
      rowWithNull.add(null);
      rowsWithNull.add(rowWithNull);
      assertThrows(NullPointerException.class, () -> sheet.setRange("A1", rowsWithNull));
    }
  }

  @Test
  void managesMergedRegionsSizingAndFreezePanes() throws Exception {
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

      assertSame(sheet, sheet.freezePanes(1, 2, 3, 4));
      assertNotNull(poiSheet.getPaneInformation());
      assertEquals(1, poiSheet.getPaneInformation().getVerticalSplitPosition());
      assertEquals(2, poiSheet.getPaneInformation().getHorizontalSplitPosition());
      assertEquals(3, poiSheet.getPaneInformation().getVerticalSplitLeftColumn());
      assertEquals(4, poiSheet.getPaneInformation().getHorizontalSplitTopRow());
      assertSame(sheet, sheet.freezePanes(0, 2, 0, 2));
      assertEquals(0, poiSheet.getPaneInformation().getVerticalSplitPosition());
      assertEquals(2, poiSheet.getPaneInformation().getHorizontalSplitPosition());
      assertEquals(0, poiSheet.getPaneInformation().getVerticalSplitLeftColumn());
      assertEquals(2, poiSheet.getPaneInformation().getHorizontalSplitTopRow());
      assertSame(sheet, sheet.freezePanes(2, 0, 2, 0));
      assertEquals(2, poiSheet.getPaneInformation().getVerticalSplitPosition());
      assertEquals(0, poiSheet.getPaneInformation().getHorizontalSplitPosition());
      assertEquals(2, poiSheet.getPaneInformation().getVerticalSplitLeftColumn());
      assertEquals(0, poiSheet.getPaneInformation().getHorizontalSplitTopRow());

      assertSame(sheet, sheet.unmergeCells("A1:B2"));
      assertEquals(0, poiSheet.getNumMergedRegions());
    }
  }

  @Test
  void validatesStructuralLayoutHelpers() throws Exception {
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
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelSheet.toRowHeightPoints((Short.MAX_VALUE / 20.0d) + 1.0d));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelSheet.toRowHeightPoints(Double.MIN_VALUE));

    assertDoesNotThrow(() -> ExcelSheet.requireFreezePaneCoordinates(1, 2, 1, 2));
    assertDoesNotThrow(() -> ExcelSheet.requireFreezePaneCoordinates(0, 2, 0, 2));
    assertDoesNotThrow(() -> ExcelSheet.requireFreezePaneCoordinates(2, 0, 2, 0));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelSheet.requireFreezePaneCoordinates(0, 0, 0, 0));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelSheet.requireFreezePaneCoordinates(0, 1, 1, 1));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelSheet.requireFreezePaneCoordinates(1, 0, 1, 1));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelSheet.requireFreezePaneCoordinates(2, 1, 1, 1));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelSheet.requireFreezePaneCoordinates(1, 2, 1, 1));

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

  @Test
  void wrapsFormulaFailuresDuringReadsAndSnapshots() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      Row row = poiSheet.createRow(0);
      row.createCell(0).setCellFormula("1+1");

      ExcelSheet writeFailureSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              poiWorkbook.getCreationHelper().createFormulaEvaluator());
      InvalidFormulaException invalidWrite =
          assertThrows(
              InvalidFormulaException.class,
              () -> writeFailureSheet.setCell("B1", ExcelCellValue.formula("SUM(")));
      assertEquals("SUM(", invalidWrite.formula());

      FormulaEvaluator baseEvaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet invalidFormulaSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              evaluateThrowingEvaluator(
                  baseEvaluator, new org.apache.poi.ss.formula.FakeFormulaFailure("bad formula")));

      InvalidFormulaException invalidSnapshot =
          assertThrows(InvalidFormulaException.class, () -> invalidFormulaSheet.snapshotCell("A1"));
      assertEquals("1+1", invalidSnapshot.formula());
      InvalidFormulaException invalidNumber =
          assertThrows(InvalidFormulaException.class, () -> invalidFormulaSheet.number("A1"));
      assertEquals("1+1", invalidNumber.formula());
      InvalidFormulaException invalidBoolean =
          assertThrows(InvalidFormulaException.class, () -> invalidFormulaSheet.bool("A1"));
      assertEquals("1+1", invalidBoolean.formula());

      ExcelSheet displayFailureSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              throwingEvaluator(
                  new org.apache.poi.ss.formula.FakeFormulaFailure("display failure")));
      InvalidFormulaException displayFailure =
          assertThrows(InvalidFormulaException.class, () -> displayFailureSheet.snapshotCell("A1"));
      assertEquals("1+1", displayFailure.formula());

      ExcelSheet unsupportedFormulaSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              throwingEvaluator(
                  new org.apache.poi.ss.formula.eval.FakeNotImplementedFunctionException(
                      "unsupported")));
      UnsupportedFormulaException unsupported =
          assertThrows(
              UnsupportedFormulaException.class, () -> unsupportedFormulaSheet.snapshotCell("A1"));
      assertEquals("1+1", unsupported.formula());

      ExcelSheet nullEvaluatedCellSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              nullEvaluatingFormulaEvaluator(baseEvaluator));
      ExcelCellSnapshot blankEvaluatedFormula = nullEvaluatedCellSheet.snapshotCell("A1");
      assertEquals("FORMULA", blankEvaluatedFormula.effectiveType());
      assertThrows(IllegalStateException.class, () -> nullEvaluatedCellSheet.number("A1"));
      assertThrows(IllegalStateException.class, () -> nullEvaluatedCellSheet.bool("A1"));
      assertInstanceOf(
          ExcelCellSnapshot.BlankSnapshot.class,
          ((ExcelCellSnapshot.FormulaSnapshot) blankEvaluatedFormula).evaluation());
    }
  }

  @Test
  void previewRowReturnsEmptyCellsForGapRows() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Sparse");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.text("first"));
      sheet.setCell("A3", ExcelCellValue.text("third"));

      List<ExcelPreviewRow> preview = sheet.preview(3, 1);
      assertEquals(3, preview.size());
      assertEquals(1, preview.get(0).cells().size());
      assertEquals(0, preview.get(1).cells().size());
      assertEquals(1, preview.get(2).cells().size());
    }
  }

  private FormulaEvaluator throwingEvaluator(RuntimeException exception) {
    InvocationHandler handler =
        new InvocationHandler() {
          @Override
          public Object invoke(Object proxy, java.lang.reflect.Method method, Object[] args)
              throws Throwable {
            if (method.getDeclaringClass() == Object.class) {
              return switch (method.getName()) {
                case "toString" -> "throwingEvaluator";
                case "hashCode" -> System.identityHashCode(proxy);
                case "equals" -> sameProxyHandler(args, this);
                default -> throw new UnsupportedOperationException(method.getName());
              };
            }
            throw exception;
          }
        };
    return (FormulaEvaluator)
        Proxy.newProxyInstance(
            formulaEvaluatorClassLoader(), new Class<?>[] {FormulaEvaluator.class}, handler);
  }

  private FormulaEvaluator evaluateThrowingEvaluator(
      FormulaEvaluator delegate, RuntimeException exception) {
    InvocationHandler handler =
        new InvocationHandler() {
          @Override
          public Object invoke(Object proxy, java.lang.reflect.Method method, Object[] args)
              throws Throwable {
            if (method.getDeclaringClass() == Object.class) {
              return switch (method.getName()) {
                case "toString" -> "evaluateThrowingEvaluator";
                case "hashCode" -> System.identityHashCode(proxy);
                case "equals" -> sameProxyHandler(args, this);
                default -> throw new UnsupportedOperationException(method.getName());
              };
            }
            if ("evaluate".equals(method.getName())) {
              throw exception;
            }
            return method.invoke(delegate, args);
          }
        };
    return (FormulaEvaluator)
        Proxy.newProxyInstance(
            formulaEvaluatorClassLoader(), new Class<?>[] {FormulaEvaluator.class}, handler);
  }

  private FormulaEvaluator nullEvaluatingFormulaEvaluator(FormulaEvaluator delegate) {
    InvocationHandler handler =
        new InvocationHandler() {
          @Override
          public Object invoke(Object proxy, java.lang.reflect.Method method, Object[] args)
              throws Throwable {
            if (method.getDeclaringClass() == Object.class) {
              return switch (method.getName()) {
                case "toString" -> "nullEvaluatingFormulaEvaluator";
                case "hashCode" -> System.identityHashCode(proxy);
                case "equals" -> sameProxyHandler(args, this);
                default -> throw new UnsupportedOperationException(method.getName());
              };
            }
            if ("evaluate".equals(method.getName())) {
              return null;
            }
            return method.invoke(delegate, args);
          }
        };
    return (FormulaEvaluator)
        Proxy.newProxyInstance(
            formulaEvaluatorClassLoader(), new Class<?>[] {FormulaEvaluator.class}, handler);
  }

  private ClassLoader formulaEvaluatorClassLoader() {
    return Objects.requireNonNull(
        Thread.currentThread().getContextClassLoader(), "context class loader must not be null");
  }

  private boolean sameProxyHandler(Object[] args, InvocationHandler handler) {
    if (args == null
        || args.length != 1
        || args[0] == null
        || !Proxy.isProxyClass(args[0].getClass())) {
      return false;
    }
    return Objects.equals(Proxy.getInvocationHandler(args[0]), handler);
  }
}
