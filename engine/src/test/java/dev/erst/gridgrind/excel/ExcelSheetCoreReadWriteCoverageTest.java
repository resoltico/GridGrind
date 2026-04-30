package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.PaneType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** ExcelSheet core read, write, parsing, and facade coverage. */
class ExcelSheetCoreReadWriteCoverageTest extends ExcelSheetTestSupport {
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
      assertNull(textSnapshot.richText());

      ExcelCellSnapshot.NumberSnapshot numberSnapshot =
          (ExcelCellSnapshot.NumberSnapshot) sheet.snapshotCell("B1");
      assertEquals("NUMBER", numberSnapshot.declaredType());
      assertEquals("NUMBER", numberSnapshot.effectiveType());
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
          () ->
              sheet.setRowHeight(0, 0, Math.nextUp(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS)));
      assertThrows(
          IllegalArgumentException.class, () -> sheet.setRowHeight(0, 0, Double.MIN_VALUE));
      assertThrows(IllegalArgumentException.class, () -> sheet.setRowHeight(0, 0, Double.NaN));
      assertThrows(NullPointerException.class, () -> sheet.setPane(null));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setPane(new ExcelSheetPane.Frozen(-1, 0, 0, 0)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setPane(new ExcelSheetPane.Frozen(0, 0, 0, 0)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setPane(new ExcelSheetPane.Frozen(0, 1, 1, 1)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setPane(new ExcelSheetPane.Frozen(1, 0, 1, 1)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setPane(new ExcelSheetPane.Frozen(2, 1, 1, 1)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setPane(new ExcelSheetPane.Frozen(1, 2, 1, 1)));
      assertThrows(IllegalArgumentException.class, () -> sheet.setZoom(9));
      assertThrows(IllegalArgumentException.class, () -> sheet.setZoom(401));
      assertThrows(NullPointerException.class, () -> sheet.setPrintLayout(null));
      assertThrows(
          NullPointerException.class,
          () -> sheet.applyStyle(null, ExcelCellStyle.numberFormat("0")));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.applyStyle(" ", ExcelCellStyle.numberFormat("0")));
      assertThrows(NullPointerException.class, () -> sheet.applyStyle("A1", null));
      assertThrows(
          NullPointerException.class,
          () -> sheet.setHyperlink(null, new ExcelHyperlink.Url("https://example.com")));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setHyperlink(" ", new ExcelHyperlink.Url("https://example.com")));
      assertThrows(NullPointerException.class, () -> sheet.setHyperlink("A1", null));
      assertThrows(NullPointerException.class, () -> sheet.clearHyperlink(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.clearHyperlink(" "));
      assertThrows(
          NullPointerException.class,
          () -> sheet.setComment(null, new ExcelComment("Review", "GridGrind", false)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.setComment(" ", new ExcelComment("Review", "GridGrind", false)));
      assertThrows(NullPointerException.class, () -> sheet.setComment("A1", null));
      assertThrows(NullPointerException.class, () -> sheet.clearComment(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.clearComment(" "));
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
      InvalidCellAddressException badAddrSnapshot =
          assertThrows(InvalidCellAddressException.class, () -> sheet.snapshotCell("BADADDR"));
      assertEquals("BADADDR", badAddrSnapshot.address());
      InvalidCellAddressException a0Snapshot =
          assertThrows(InvalidCellAddressException.class, () -> sheet.snapshotCell("A0"));
      assertEquals("A0", a0Snapshot.address());
      InvalidCellAddressException numericOnlySnapshot =
          assertThrows(InvalidCellAddressException.class, () -> sheet.snapshotCell("1"));
      assertEquals("1", numericOnlySnapshot.address());
      InvalidCellAddressException outOfBoundsRow =
          assertThrows(InvalidCellAddressException.class, () -> sheet.snapshotCell("A1048577"));
      assertEquals("A1048577", outOfBoundsRow.address());
      InvalidCellAddressException outOfBoundsCol =
          assertThrows(InvalidCellAddressException.class, () -> sheet.snapshotCell("XFE1"));
      assertEquals("XFE1", outOfBoundsCol.address());
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
  void groupsAndUngroupsRowsAndColumnsThroughSheetFacade() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.text("Header"));

      assertSame(sheet, sheet.groupRows(new ExcelRowSpan(1, 3), true));
      assertSame(sheet, sheet.groupColumns(new ExcelColumnSpan(1, 3), true));
      assertSame(sheet, sheet.ungroupRows(new ExcelRowSpan(1, 3)));
      assertSame(sheet, sheet.ungroupColumns(new ExcelColumnSpan(1, 3)));

      WorkbookSheetResult.SheetLayout layout = sheet.layout();
      assertFalse(layout.rows().get(1).hidden());
      assertEquals(0, layout.rows().get(1).outlineLevel());
      assertFalse(layout.columns().get(1).hidden());
      assertEquals(0, layout.columns().get(1).outlineLevel());
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
              new ExcelCellAlignment(
                  true, ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.TOP, null, null),
              new ExcelCellFont(true, null, null, null, null, null, null),
              null,
              null,
              null));
      sheet.applyStyle("C1", ExcelCellStyle.emphasis(null, true));

      ExcelCellSnapshot styledValue = sheet.snapshotCell("A1");
      assertEquals("#,##0.00", styledValue.style().numberFormat());
      assertTrue(styledValue.style().font().bold());
      assertTrue(styledValue.style().alignment().wrapText());
      assertEquals(
          ExcelHorizontalAlignment.CENTER, styledValue.style().alignment().horizontalAlignment());
      assertEquals(ExcelVerticalAlignment.TOP, styledValue.style().alignment().verticalAlignment());
      assertEquals("Calibri", styledValue.style().font().fontName());
      assertEquals(220, styledValue.style().font().fontHeight().twips());
      assertFalse(styledValue.style().font().underline());
      assertFalse(styledValue.style().font().strikeout());
      assertEquals(ExcelColorSnapshot.indexed(8), styledValue.style().font().fontColor());
      assertNull(fillForegroundColor(styledValue.style().fill()));
      assertEquals(ExcelBorderStyle.NONE, styledValue.style().border().top().style());
      assertEquals(ExcelBorderStyle.NONE, styledValue.style().border().right().style());
      assertEquals(ExcelBorderStyle.NONE, styledValue.style().border().bottom().style());
      assertEquals(ExcelBorderStyle.NONE, styledValue.style().border().left().style());

      List<ExcelPreviewRow> preview = sheet.preview(2, 3);
      assertTrue(preview.getFirst().cells().stream().anyMatch(cell -> "C1".equals(cell.address())));
      assertEquals("BLANK", sheet.snapshotCell("C1").effectiveType());
      assertTrue(sheet.snapshotCell("C1").style().font().italic());

      sheet.clearRange("A2:B2");

      ExcelCellSnapshot cleared = sheet.snapshotCell("A2");
      assertEquals("BLANK", cleared.declaredType());
      assertEquals("General", cleared.style().numberFormat());
      assertFalse(cleared.style().font().bold());
      assertEquals(
          ExcelHorizontalAlignment.GENERAL, cleared.style().alignment().horizontalAlignment());
      assertEquals(ExcelVerticalAlignment.BOTTOM, cleared.style().alignment().verticalAlignment());
      assertEquals("Calibri", cleared.style().font().fontName());
      assertEquals(220, cleared.style().font().fontHeight().twips());
      assertEquals(ExcelColorSnapshot.indexed(8), cleared.style().font().fontColor());
      assertFalse(cleared.style().font().underline());
      assertFalse(cleared.style().font().strikeout());
      assertNull(fillForegroundColor(cleared.style().fill()));
      assertEquals(ExcelBorderStyle.NONE, cleared.style().border().top().style());
      assertEquals(ExcelBorderStyle.NONE, cleared.style().border().right().style());
      assertEquals(ExcelBorderStyle.NONE, cleared.style().border().bottom().style());
      assertEquals(ExcelBorderStyle.NONE, cleared.style().border().left().style());

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
  void mergesFormattingDepthStylesAndPreservesExistingAttributes() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.text("Item"));
      sheet.applyStyle(
          "A1",
          new ExcelCellStyle(
              null,
              new ExcelCellAlignment(
                  true, ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.TOP, null, null),
              new ExcelCellFont(
                  true,
                  false,
                  "Aptos",
                  new ExcelFontHeight(280),
                  ExcelColor.rgb("#1F4E78"),
                  true,
                  false),
              ExcelCellFill.patternForeground(ExcelFillPattern.SOLID, ExcelColor.rgb("#FFF2CC")),
              new ExcelBorder(new ExcelBorderSide(ExcelBorderStyle.THIN), null, null, null, null),
              null));
      sheet.applyStyle(
          "A1",
          new ExcelCellStyle(
              null,
              null,
              new ExcelCellFont(null, null, null, null, null, null, true),
              null,
              new ExcelBorder(null, null, new ExcelBorderSide(ExcelBorderStyle.DOUBLE), null, null),
              null));

      ExcelCellSnapshot styled = sheet.snapshotCell("A1");
      assertTrue(styled.style().font().bold());
      assertFalse(styled.style().font().italic());
      assertTrue(styled.style().alignment().wrapText());
      assertEquals(
          ExcelHorizontalAlignment.CENTER, styled.style().alignment().horizontalAlignment());
      assertEquals(ExcelVerticalAlignment.TOP, styled.style().alignment().verticalAlignment());
      assertEquals("Aptos", styled.style().font().fontName());
      assertEquals(280, styled.style().font().fontHeight().twips());
      assertEquals(rgb("#1F4E78"), styled.style().font().fontColor());
      assertTrue(styled.style().font().underline());
      assertTrue(styled.style().font().strikeout());
      assertEquals(rgb("#FFF2CC"), fillForegroundColor(styled.style().fill()));
      assertEquals(ExcelBorderStyle.THIN, styled.style().border().top().style());
      assertEquals(ExcelBorderStyle.DOUBLE, styled.style().border().right().style());
      assertEquals(ExcelBorderStyle.THIN, styled.style().border().bottom().style());
      assertEquals(ExcelBorderStyle.THIN, styled.style().border().left().style());
    }
  }

  @Test
  void snapshotsAndClearsHyperlinksAndComments() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
      sheet.setComment("A1", new ExcelComment("Review", "GridGrind", true));

      ExcelCellSnapshot.BlankSnapshot snapshot =
          (ExcelCellSnapshot.BlankSnapshot) sheet.snapshotCell("A1");
      assertEquals(
          new ExcelHyperlink.Url("https://example.com/report"),
          snapshot.metadata().hyperlink().orElseThrow());
      assertEquals(
          new ExcelComment("Review", "GridGrind", true),
          snapshot.metadata().comment().orElseThrow().toPlainComment());

      List<ExcelPreviewRow> preview = sheet.preview(1, 1);
      assertEquals(1, preview.size());
      assertEquals("A1", preview.getFirst().cells().getFirst().address());

      sheet.clearHyperlink("A1");
      sheet.clearComment("A1");
      ExcelCellSnapshot.BlankSnapshot clearedMetadata =
          (ExcelCellSnapshot.BlankSnapshot) sheet.snapshotCell("A1");
      assertTrue(clearedMetadata.metadata().hyperlink().isEmpty());
      assertTrue(clearedMetadata.metadata().comment().isEmpty());

      sheet.setHyperlink("A1", new ExcelHyperlink.Document("Budget!B4"));
      sheet.setComment("A1", new ExcelComment("Again", "GridGrind", false));
      sheet.clearRange("A1");
      ExcelCellSnapshot.BlankSnapshot clearedRange =
          (ExcelCellSnapshot.BlankSnapshot) sheet.snapshotCell("A1");
      assertTrue(clearedRange.metadata().hyperlink().isEmpty());
      assertTrue(clearedRange.metadata().comment().isEmpty());

      // clearHyperlink and clearComment are no-ops on cells that do not physically exist
      assertDoesNotThrow(() -> sheet.clearHyperlink("B2"));
      assertDoesNotThrow(() -> sheet.clearComment("B2"));
      // calling again on B2 (still non-existent) must still be a no-op, not throw
      assertDoesNotThrow(() -> sheet.clearHyperlink("B2"));
      assertDoesNotThrow(() -> sheet.clearComment("B2"));
    }
  }

  @Test
  void writesAndReadsFileHyperlinksWithSpacesInPlainPaths() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertDoesNotThrow(
          () -> sheet.setHyperlink("A1", new ExcelHyperlink.File("support/budget backup.xlsx")));

      Cell cell = poiSheet.getRow(0).getCell(0);
      assertEquals("support/budget%20backup.xlsx", cell.getHyperlink().getAddress());
      assertEquals(
          new ExcelHyperlink.File("support/budget backup.xlsx"), ExcelSheet.hyperlink(cell));
    }
  }

  @Test
  void clearRangeIsNoOpOnNeverWrittenCells() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Data");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.text("anchor"));

      int rowsBefore = sheet.physicalRowCount();
      int lastRowBefore = sheet.lastRowIndex();
      int lastColBefore = sheet.lastColumnIndex();

      // clear a range that has never been written
      sheet.clearRange("B2:E5");

      // physicalRowCount, lastRowIndex, and lastColumnIndex must not change
      assertEquals(rowsBefore, sheet.physicalRowCount(), "physicalRowCount must not increase");
      assertEquals(lastRowBefore, sheet.lastRowIndex(), "lastRowIndex must not change");
      assertEquals(lastColBefore, sheet.lastColumnIndex(), "lastColumnIndex must not change");
    }
  }

  @Test
  void clearRangeSkipsAbsentCellsInExistingRows() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Data");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      // Row 2 exists (B2 is written) but C2 has never been written.
      // Clearing B2:C2 must not throw even though C2 is absent.
      sheet.setCell("B2", ExcelCellValue.text("present"));
      assertDoesNotThrow(() -> sheet.clearRange("B2:C2"));
      // B2 must now be blank after the clear
      ExcelCellSnapshot b2 = sheet.snapshotCell("B2");
      assertEquals("BLANK", b2.effectiveType());
    }
  }

  @Test
  void introspectsWindowsSelectionsAndLayoutAcrossSparseMetadata() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      WorkbookSheetResult.SheetLayout emptyLayout = sheet.layout();
      assertEquals(new ExcelSheetPane.None(), emptyLayout.pane());
      assertEquals(100, emptyLayout.zoomPercent());
      assertEquals(ExcelSheetDisplay.defaults(), emptyLayout.presentation().display());
      assertEquals(
          ExcelSheetOutlineSummary.defaults(), emptyLayout.presentation().outlineSummary());
      assertEquals(ExcelSheetDefaults.defaults(), emptyLayout.presentation().sheetDefaults());
      assertEquals(List.of(), emptyLayout.columns());
      assertEquals(List.of(), emptyLayout.rows());

      sheet.setCell("B2", ExcelCellValue.text("Center"));
      sheet.setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
      sheet.setComment("C3", new ExcelComment("Review", "GridGrind", false));
      sheet.setColumnWidth(0, 0, 12.5);
      sheet.setRowHeight(0, 0, 19.5);
      sheet.setPresentation(
          new ExcelSheetPresentation(
              new ExcelSheetDisplay(false, false, false, true, true),
              ExcelColor.rgb("#112233"),
              new ExcelSheetOutlineSummary(false, false),
              new ExcelSheetDefaults(11, 18.5d),
              List.of(
                  new ExcelIgnoredError(
                      "A1:B2", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)))));

      WorkbookSheetResult.Window window = sheet.window("A1", 3, 3);
      assertEquals("A1", window.rows().getFirst().cells().getFirst().address());
      assertEquals("B2", window.rows().get(1).cells().get(1).address());
      assertEquals("Center", window.rows().get(1).cells().get(1).displayValue());
      assertEquals("C3", window.rows().get(2).cells().get(2).address());
      assertThrows(IllegalArgumentException.class, () -> sheet.window("A1", 0, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.window("A1", 1, 0));
      assertThrows(IllegalArgumentException.class, () -> sheet.window("A1", 1048577, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.window("A1", 1, 16385));

      List<WorkbookSheetResult.CellHyperlink> allHyperlinks =
          sheet.hyperlinks(new ExcelCellSelection.AllUsedCells());
      assertEquals(1, allHyperlinks.size());
      assertEquals("A1", allHyperlinks.getFirst().address());

      List<WorkbookSheetResult.CellHyperlink> selectedHyperlinks =
          sheet.hyperlinks(new ExcelCellSelection.Selected(List.of("A1", "B2", "B9", "D3")));
      assertEquals(1, selectedHyperlinks.size());
      assertEquals("A1", selectedHyperlinks.getFirst().address());

      List<WorkbookSheetResult.CellComment> selectedComments =
          sheet.comments(new ExcelCellSelection.Selected(List.of("C3", "B2", "A9", "D3")));
      assertEquals(1, selectedComments.size());
      assertEquals("C3", selectedComments.getFirst().address());

      WorkbookSheetResult.SheetLayout unfrozenLayout = sheet.layout();
      assertEquals(new ExcelSheetPane.None(), unfrozenLayout.pane());
      assertEquals(100, unfrozenLayout.zoomPercent());
      assertEquals(
          new ExcelSheetDisplay(false, false, false, true, true),
          unfrozenLayout.presentation().display());
      assertEquals(ExcelColorSnapshot.rgb("#112233"), unfrozenLayout.presentation().tabColor());
      assertEquals(
          new ExcelSheetOutlineSummary(false, false),
          unfrozenLayout.presentation().outlineSummary());
      assertEquals(
          new ExcelSheetDefaults(11, 18.5d), unfrozenLayout.presentation().sheetDefaults());
      assertEquals(
          List.of(
              new ExcelIgnoredError("A1:B2", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))),
          unfrozenLayout.presentation().ignoredErrors());
      assertEquals(3, unfrozenLayout.columns().size());
      assertEquals(3, unfrozenLayout.rows().size());

      sheet.setPane(new ExcelSheetPane.Frozen(1, 1, 1, 1));
      sheet.setZoom(135);
      WorkbookSheetResult.SheetLayout frozenLayout = sheet.layout();
      assertEquals(new ExcelSheetPane.Frozen(1, 1, 1, 1), frozenLayout.pane());
      assertEquals(135, frozenLayout.zoomPercent());

      poiSheet.createSplitPane(2000, 2000, 0, 0, PaneType.LOWER_RIGHT);
      WorkbookSheetResult.SheetLayout splitLayout = sheet.layout();
      assertEquals(
          new ExcelSheetPane.Split(2000, 2000, 0, 0, ExcelPaneRegion.LOWER_RIGHT),
          splitLayout.pane());
    }
  }

  @Test
  void replacesHyperlinksOnRepeatedWrites() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setHyperlink("F18", new ExcelHyperlink.Email("Report_Value@example.com"));
      sheet.setHyperlink("F18", new ExcelHyperlink.Email("Summary.Total@example.com"));

      ExcelCellSnapshot.BlankSnapshot snapshot =
          (ExcelCellSnapshot.BlankSnapshot) sheet.snapshotCell("F18");
      assertEquals(
          new ExcelHyperlink.Email("Summary.Total@example.com"),
          snapshot.metadata().hyperlink().orElseThrow());
    }
  }
}
