package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.PaneType;
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
          () -> sheet.setRowHeight(0, 0, (Short.MAX_VALUE / 20.0d) + 1.0d));
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

      WorkbookReadResult.SheetLayout layout = sheet.layout();
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
      assertEquals("#000000", styledValue.style().font().fontColor());
      assertNull(styledValue.style().fill().foregroundColor());
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
      assertEquals("#000000", cleared.style().font().fontColor());
      assertFalse(cleared.style().font().underline());
      assertFalse(cleared.style().font().strikeout());
      assertNull(cleared.style().fill().foregroundColor());
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
                  true, false, "Aptos", new ExcelFontHeight(280), "#1F4E78", true, false),
              new ExcelCellFill(ExcelFillPattern.SOLID, "#FFF2CC", null),
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
      assertEquals("#1F4E78", styled.style().font().fontColor());
      assertTrue(styled.style().font().underline());
      assertTrue(styled.style().font().strikeout());
      assertEquals("#FFF2CC", styled.style().fill().foregroundColor());
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
          snapshot.metadata().comment().orElseThrow());

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

      WorkbookReadResult.SheetLayout emptyLayout = sheet.layout();
      assertEquals(new ExcelSheetPane.None(), emptyLayout.pane());
      assertEquals(100, emptyLayout.zoomPercent());
      assertEquals(List.of(), emptyLayout.columns());
      assertEquals(List.of(), emptyLayout.rows());

      sheet.setCell("B2", ExcelCellValue.text("Center"));
      sheet.setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
      sheet.setComment("C3", new ExcelComment("Review", "GridGrind", false));
      sheet.setColumnWidth(0, 0, 12.5);
      sheet.setRowHeight(0, 0, 19.5);

      WorkbookReadResult.Window window = sheet.window("A1", 3, 3);
      assertEquals("A1", window.rows().getFirst().cells().getFirst().address());
      assertEquals("B2", window.rows().get(1).cells().get(1).address());
      assertEquals("Center", window.rows().get(1).cells().get(1).displayValue());
      assertEquals("C3", window.rows().get(2).cells().get(2).address());
      assertThrows(IllegalArgumentException.class, () -> sheet.window("A1", 0, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.window("A1", 1, 0));
      assertThrows(IllegalArgumentException.class, () -> sheet.window("A1", 1048577, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.window("A1", 1, 16385));

      List<WorkbookReadResult.CellHyperlink> allHyperlinks =
          sheet.hyperlinks(new ExcelCellSelection.AllUsedCells());
      assertEquals(1, allHyperlinks.size());
      assertEquals("A1", allHyperlinks.getFirst().address());

      List<WorkbookReadResult.CellHyperlink> selectedHyperlinks =
          sheet.hyperlinks(new ExcelCellSelection.Selected(List.of("A1", "B2", "B9", "D3")));
      assertEquals(1, selectedHyperlinks.size());
      assertEquals("A1", selectedHyperlinks.getFirst().address());

      List<WorkbookReadResult.CellComment> selectedComments =
          sheet.comments(new ExcelCellSelection.Selected(List.of("C3", "B2", "A9", "D3")));
      assertEquals(1, selectedComments.size());
      assertEquals("C3", selectedComments.getFirst().address());

      WorkbookReadResult.SheetLayout unfrozenLayout = sheet.layout();
      assertEquals(new ExcelSheetPane.None(), unfrozenLayout.pane());
      assertEquals(100, unfrozenLayout.zoomPercent());
      assertEquals(3, unfrozenLayout.columns().size());
      assertEquals(3, unfrozenLayout.rows().size());

      sheet.setPane(new ExcelSheetPane.Frozen(1, 1, 1, 1));
      sheet.setZoom(135);
      WorkbookReadResult.SheetLayout frozenLayout = sheet.layout();
      assertEquals(new ExcelSheetPane.Frozen(1, 1, 1, 1), frozenLayout.pane());
      assertEquals(135, frozenLayout.zoomPercent());

      poiSheet.createSplitPane(2000, 2000, 0, 0, PaneType.LOWER_RIGHT);
      WorkbookReadResult.SheetLayout splitLayout = sheet.layout();
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

  @Test
  void appendRowIgnoresMetadataOnlyRowsWhenChoosingTheNextDataRow() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.text("Header"));
      sheet.setHyperlink("A4", new ExcelHyperlink.Document("Budget!A1"));
      sheet.setComment("B5", new ExcelComment("Note", "GridGrind", false));
      sheet.appendRow(ExcelCellValue.text("Hosting"), ExcelCellValue.number(49.0));

      assertEquals("Hosting", sheet.text("A2"));
      assertEquals(49.0, sheet.number("B2"));
      assertEquals(
          new ExcelHyperlink.Document("Budget!A1"),
          sheet.snapshotCell("A4").metadata().hyperlink().orElseThrow());
      assertEquals(
          new ExcelComment("Note", "GridGrind", false),
          sheet.snapshotCell("B5").metadata().comment().orElseThrow());
    }
  }

  @Test
  void setCellPreservesExistingStyleAndMetadataOnStyledBlankCells() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.applyStyle(
          "A1",
          new ExcelCellStyle(
              null,
              new ExcelCellAlignment(null, ExcelHorizontalAlignment.CENTER, null, null, null),
              new ExcelCellFont(Boolean.TRUE, null, null, null, null, null, null),
              new ExcelCellFill(ExcelFillPattern.SOLID, "#AABBCC", null),
              new ExcelBorder(new ExcelBorderSide(ExcelBorderStyle.THIN), null, null, null, null),
              null));
      sheet.setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
      sheet.setComment("A1", new ExcelComment("Review", "GridGrind", false));

      ExcelCellSnapshot.TextSnapshot textSnapshot =
          (ExcelCellSnapshot.TextSnapshot)
              sheet.setCell("A1", ExcelCellValue.text("Quarterly report")).snapshotCell("A1");
      assertEquals("Quarterly report", textSnapshot.stringValue());
      assertEquals("#AABBCC", textSnapshot.style().fill().foregroundColor());
      assertEquals(ExcelBorderStyle.THIN, textSnapshot.style().border().top().style());
      assertEquals(
          ExcelHorizontalAlignment.CENTER, textSnapshot.style().alignment().horizontalAlignment());
      assertEquals(
          new ExcelHyperlink.Url("https://example.com/report"),
          textSnapshot.metadata().hyperlink().orElseThrow());
      assertEquals(
          new ExcelComment("Review", "GridGrind", false),
          textSnapshot.metadata().comment().orElseThrow());

      ExcelCellSnapshot.BlankSnapshot blankSnapshot =
          (ExcelCellSnapshot.BlankSnapshot)
              sheet.setCell("A1", ExcelCellValue.blank()).snapshotCell("A1");
      assertEquals("#AABBCC", blankSnapshot.style().fill().foregroundColor());
      assertEquals(ExcelBorderStyle.THIN, blankSnapshot.style().border().top().style());
      assertEquals(
          new ExcelHyperlink.Url("https://example.com/report"),
          blankSnapshot.metadata().hyperlink().orElseThrow());
      assertEquals(
          new ExcelComment("Review", "GridGrind", false),
          blankSnapshot.metadata().comment().orElseThrow());
    }
  }

  @Test
  void dateValueWritesMergeRequiredNumberFormatsOntoExistingStyle() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      ExcelCellStyle style =
          new ExcelCellStyle(
              null,
              new ExcelCellAlignment(Boolean.TRUE, null, ExcelVerticalAlignment.TOP, null, null),
              new ExcelCellFont(Boolean.TRUE, null, null, null, null, null, null),
              new ExcelCellFill(ExcelFillPattern.SOLID, "#DDEBF7", null),
              new ExcelBorder(null, null, new ExcelBorderSide(ExcelBorderStyle.DOUBLE), null, null),
              null);
      sheet.applyStyle("A1:B1", style);

      ExcelCellSnapshot.NumberSnapshot dateSnapshot =
          (ExcelCellSnapshot.NumberSnapshot)
              sheet
                  .setCell("A1", ExcelCellValue.date(LocalDate.of(2026, 3, 31)))
                  .snapshotCell("A1");
      ExcelCellSnapshot.NumberSnapshot dateTimeSnapshot =
          (ExcelCellSnapshot.NumberSnapshot)
              sheet
                  .setCell("B1", ExcelCellValue.dateTime(LocalDateTime.of(2026, 3, 31, 9, 45, 0)))
                  .snapshotCell("B1");

      assertEquals("yyyy-mm-dd", dateSnapshot.style().numberFormat());
      assertEquals("#DDEBF7", dateSnapshot.style().fill().foregroundColor());
      assertTrue(dateSnapshot.style().font().bold());
      assertTrue(dateSnapshot.style().alignment().wrapText());
      assertEquals(
          ExcelVerticalAlignment.TOP, dateSnapshot.style().alignment().verticalAlignment());
      assertEquals(ExcelBorderStyle.DOUBLE, dateSnapshot.style().border().right().style());

      assertEquals("yyyy-mm-dd hh:mm:ss", dateTimeSnapshot.style().numberFormat());
      assertEquals("#DDEBF7", dateTimeSnapshot.style().fill().foregroundColor());
      assertTrue(dateTimeSnapshot.style().font().bold());
      assertTrue(dateTimeSnapshot.style().alignment().wrapText());
      assertEquals(
          ExcelVerticalAlignment.TOP, dateTimeSnapshot.style().alignment().verticalAlignment());
      assertEquals(ExcelBorderStyle.DOUBLE, dateTimeSnapshot.style().border().right().style());
    }
  }

  @Test
  void snapshotsRichTextRunsWithEffectiveInheritedFontFacts() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell(
          "A1",
          ExcelCellValue.richText(
              new ExcelRichText(
                  List.of(
                      new ExcelRichTextRun("Quarterly", null),
                      new ExcelRichTextRun(
                          " Report",
                          new ExcelCellFont(
                              Boolean.TRUE, null, null, null, "#FF0000", null, null))))));
      sheet.applyStyle(
          "A1",
          new ExcelCellStyle(
              null,
              null,
              new ExcelCellFont(
                  null, Boolean.TRUE, "Aptos", new ExcelFontHeight(260), "#112233", null, null),
              null,
              null,
              null));

      ExcelCellSnapshot.TextSnapshot textSnapshot =
          (ExcelCellSnapshot.TextSnapshot) sheet.snapshotCell("A1");

      assertEquals("Quarterly Report", textSnapshot.stringValue());
      assertNotNull(textSnapshot.richText());
      assertEquals("Quarterly Report", textSnapshot.richText().plainText());
      assertEquals(2, textSnapshot.richText().runs().size());

      ExcelRichTextRunSnapshot firstRun = textSnapshot.richText().runs().get(0);
      assertEquals("Quarterly", firstRun.text());
      assertEquals("Aptos", firstRun.font().fontName());
      assertEquals(new ExcelFontHeight(260), firstRun.font().fontHeight());
      assertEquals("#112233", firstRun.font().fontColor());
      assertFalse(firstRun.font().bold());
      assertTrue(firstRun.font().italic());

      ExcelRichTextRunSnapshot secondRun = textSnapshot.richText().runs().get(1);
      assertEquals(" Report", secondRun.text());
      assertEquals("Aptos", secondRun.font().fontName());
      assertEquals(new ExcelFontHeight(260), secondRun.font().fontHeight());
      assertEquals("#FF0000", secondRun.font().fontColor());
      assertTrue(secondRun.font().bold());
      assertTrue(secondRun.font().italic());
    }
  }

  @Test
  void appendRowPreservesStyleWhenReusingStyledBlankRows() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.applyStyle(
          "A1:B2",
          new ExcelCellStyle(
              null,
              null,
              new ExcelCellFont(
                  null,
                  null,
                  "Aptos",
                  ExcelFontHeight.fromPoints(new java.math.BigDecimal("15.2")),
                  "#A3A3A3",
                  null,
                  Boolean.TRUE),
              new ExcelCellFill(ExcelFillPattern.SOLID, "#CDCDCD", null),
              new ExcelBorder(
                  new ExcelBorderSide(ExcelBorderStyle.DASH_DOT), null, null, null, null),
              null));

      sheet.appendRow(ExcelCellValue.number(607.8483822864587), ExcelCellValue.bool(false));

      ExcelCellSnapshot.NumberSnapshot firstCell =
          (ExcelCellSnapshot.NumberSnapshot) sheet.snapshotCell("A1");
      ExcelCellSnapshot.BooleanSnapshot secondCell =
          (ExcelCellSnapshot.BooleanSnapshot) sheet.snapshotCell("B1");
      ExcelCellSnapshot.BlankSnapshot untouchedStyledCell =
          (ExcelCellSnapshot.BlankSnapshot) sheet.snapshotCell("A2");

      assertEquals("#CDCDCD", firstCell.style().fill().foregroundColor());
      assertEquals("#A3A3A3", firstCell.style().font().fontColor());
      assertTrue(firstCell.style().font().strikeout());
      assertEquals(ExcelBorderStyle.DASH_DOT, firstCell.style().border().top().style());
      assertEquals("#CDCDCD", secondCell.style().fill().foregroundColor());
      assertEquals(ExcelBorderStyle.DASH_DOT, secondCell.style().border().top().style());
      assertEquals("#CDCDCD", untouchedStyledCell.style().fill().foregroundColor());
      assertEquals(ExcelBorderStyle.DASH_DOT, untouchedStyledCell.style().border().top().style());
    }
  }

  @Test
  void appendRowDateTimeWriteLayersRequiredNumberFormatWhenReusingStyledBlankRows()
      throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.appendRow(
          ExcelCellValue.dateTime(LocalDateTime.of(2026, 2, 6, 13, 1, 1)),
          ExcelCellValue.dateTime(LocalDateTime.of(2026, 2, 6, 13, 58, 1)));
      sheet.applyStyle(
          "A1:B2",
          new ExcelCellStyle(
              "0.00",
              new ExcelCellAlignment(Boolean.TRUE, null, ExcelVerticalAlignment.CENTER, null, null),
              new ExcelCellFont(Boolean.TRUE, null, null, null, null, null, null),
              new ExcelCellFill(ExcelFillPattern.SOLID, "#603A79", null),
              new ExcelBorder(
                  new ExcelBorderSide(ExcelBorderStyle.MEDIUM_DASHED),
                  null,
                  new ExcelBorderSide(ExcelBorderStyle.MEDIUM_DASHED),
                  null,
                  null),
              null));

      sheet.appendRow(
          ExcelCellValue.dateTime(LocalDateTime.of(2026, 2, 6, 13, 1, 58)),
          ExcelCellValue.dateTime(LocalDateTime.of(2026, 2, 6, 13, 1, 1)));

      ExcelCellSnapshot.NumberSnapshot appendedFirstCell =
          (ExcelCellSnapshot.NumberSnapshot) sheet.snapshotCell("A2");
      ExcelCellSnapshot.NumberSnapshot appendedSecondCell =
          (ExcelCellSnapshot.NumberSnapshot) sheet.snapshotCell("B2");

      assertEquals("yyyy-mm-dd hh:mm:ss", appendedFirstCell.style().numberFormat());
      assertEquals("yyyy-mm-dd hh:mm:ss", appendedSecondCell.style().numberFormat());
      assertEquals("#603A79", appendedFirstCell.style().fill().foregroundColor());
      assertTrue(appendedFirstCell.style().font().bold());
      assertTrue(appendedFirstCell.style().alignment().wrapText());
      assertEquals(
          ExcelVerticalAlignment.CENTER, appendedFirstCell.style().alignment().verticalAlignment());
      assertEquals(
          ExcelBorderStyle.MEDIUM_DASHED, appendedFirstCell.style().border().top().style());
      assertEquals(
          ExcelBorderStyle.MEDIUM_DASHED, appendedFirstCell.style().border().right().style());
    }
  }

  @Test
  void appendRowWithNoValuesIsANoOp() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertSame(sheet, sheet.appendRow());
      assertEquals(0, sheet.physicalRowCount());
      assertEquals(-1, sheet.lastRowIndex());
      assertEquals(-1, sheet.lastColumnIndex());
    }
  }

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

      WorkbookReadResult.SheetLayout layout = sheet.layout();
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
      sheet.setRowHeight(0, 0, 1638.35);
      assertEquals(32767 / 20.0, sheet.layout().rows().get(0).heightPoints());

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
    assertDoesNotThrow(() -> ExcelSheet.toRowHeightPoints(Short.MAX_VALUE / 20.0d));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelSheet.toRowHeightPoints(Math.nextUp(Short.MAX_VALUE / 20.0d)));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelSheet.toRowHeightPoints((Short.MAX_VALUE / 20.0d) + 1.0d));
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
    assertEquals(ExcelPaneRegion.UPPER_LEFT, ExcelPaneRegion.fromPoi(PaneType.UPPER_LEFT));
    assertEquals(ExcelPaneRegion.UPPER_RIGHT, ExcelPaneRegion.fromPoi(PaneType.UPPER_RIGHT));
    assertEquals(ExcelPaneRegion.LOWER_LEFT, ExcelPaneRegion.fromPoi(PaneType.LOWER_LEFT));
    assertEquals(ExcelPaneRegion.LOWER_RIGHT, ExcelPaneRegion.fromPoi(PaneType.LOWER_RIGHT));
    assertEquals(PaneType.UPPER_LEFT, ExcelPaneRegion.UPPER_LEFT.toPoi());
    assertEquals(PaneType.UPPER_RIGHT, ExcelPaneRegion.UPPER_RIGHT.toPoi());
    assertEquals(PaneType.LOWER_LEFT, ExcelPaneRegion.LOWER_LEFT.toPoi());
    assertEquals(PaneType.LOWER_RIGHT, ExcelPaneRegion.LOWER_RIGHT.toPoi());
    assertThrows(IllegalArgumentException.class, () -> ExcelSheetViewSupport.requireZoomPercent(9));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelSheetViewSupport.requireZoomPercent(401));
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
      InvalidFormulaException parserStateWrite =
          assertThrows(
              InvalidFormulaException.class,
              () -> writeFailureSheet.setCell("C1", ExcelCellValue.formula("[^owe_e`ffffff")));
      assertEquals("[^owe_e`ffffff", parserStateWrite.formula());

      FormulaEvaluator baseEvaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet invalidFormulaSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.failingEvaluation(
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
              FormulaRuntimeTestDouble.alwaysFail(
                  new org.apache.poi.ss.formula.FakeFormulaFailure("display failure")));
      InvalidFormulaException displayFailure =
          assertThrows(InvalidFormulaException.class, () -> displayFailureSheet.snapshotCell("A1"));
      assertEquals("1+1", displayFailure.formula());

      ExcelSheet unsupportedFormulaSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.alwaysFail(
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
              FormulaRuntimeTestDouble.nullEvaluation(baseEvaluator));
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

  @Test
  void interpretsHyperlinksCommentsAndPreviewHelpersAcrossAllVariants() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      Row row = poiSheet.createRow(0);
      Cell blankCell = row.createCell(0);
      Cell urlCell = row.createCell(1);
      Cell emailCell = row.createCell(2);
      Cell fileCell = row.createCell(3);
      Cell documentCell = row.createCell(4);
      Cell commentCell = row.createCell(5);
      Cell missingStringCommentCell = row.createCell(6);
      Cell blankCommentCell = row.createCell(7);

      urlCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.URL, "https://example.com/report"));
      emailCell.setHyperlink(
          hyperlink(poiWorkbook, HyperlinkType.EMAIL, "mailto:team@example.com"));
      fileCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.FILE, "/tmp/report.xlsx"));
      documentCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.DOCUMENT, "Budget!B4"));
      commentCell.setCellComment(comment(poiWorkbook, poiSheet, "Review", "GridGrind", false));
      missingStringCommentCell.setCellComment(emptyComment(poiWorkbook, poiSheet));
      blankCommentCell.setCellComment(comment(poiWorkbook, poiSheet, " ", " ", false));

      assertFalse(ExcelSheet.shouldPreview(null));
      assertFalse(ExcelSheet.shouldPreview(blankCell));
      assertTrue(ExcelSheet.shouldPreview(urlCell));
      assertTrue(ExcelSheet.shouldPreview(commentCell));

      assertNull(ExcelSheet.hyperlink(blankCell));
      assertEquals(
          new ExcelHyperlink.Url("https://example.com/report"), ExcelSheet.hyperlink(urlCell));
      assertEquals(new ExcelHyperlink.Email("team@example.com"), ExcelSheet.hyperlink(emailCell));
      assertEquals(new ExcelHyperlink.File("/tmp/report.xlsx"), ExcelSheet.hyperlink(fileCell));
      assertEquals(new ExcelHyperlink.Document("Budget!B4"), ExcelSheet.hyperlink(documentCell));
      assertNull(ExcelSheet.hyperlink(HyperlinkType.NONE, "ignored"));
      assertNull(ExcelSheet.hyperlink((HyperlinkType) null, "ignored"));
      assertNull(ExcelSheet.hyperlink(HyperlinkType.URL, null));
      assertNull(ExcelSheet.hyperlink(HyperlinkType.URL, " "));
      assertNull(ExcelSheet.hyperlink(HyperlinkType.EMAIL, "mailto:"));
      assertNull(ExcelSheet.hyperlink(HyperlinkType.FILE, "https://example.com/report"));

      assertEquals(new ExcelComment("Review", "GridGrind", false), ExcelSheet.comment(commentCell));
      assertNull(ExcelSheet.comment(blankCell));
      assertNull(ExcelSheet.comment(missingStringCommentCell));
      assertNull(ExcelSheet.comment(blankCommentCell));
      Cell nullAuthorCommentCell = row.createCell(11);
      nullAuthorCommentCell.setCellComment(comment(poiWorkbook, poiSheet, "Review", null, false));
      assertNull(ExcelSheet.comment(nullAuthorCommentCell));
      assertNull(ExcelSheet.comment((String) null, "GridGrind", false));
      Cell blankAuthorCommentCell = row.createCell(12);
      blankAuthorCommentCell.setCellComment(comment(poiWorkbook, poiSheet, "Review", " ", false));
      assertNull(ExcelSheet.comment(blankAuthorCommentCell));

      assertEquals(HyperlinkType.URL, ExcelSheet.toPoi(ExcelHyperlinkType.URL));
      assertEquals(HyperlinkType.EMAIL, ExcelSheet.toPoi(ExcelHyperlinkType.EMAIL));
      assertEquals(HyperlinkType.FILE, ExcelSheet.toPoi(ExcelHyperlinkType.FILE));
      assertEquals(HyperlinkType.DOCUMENT, ExcelSheet.toPoi(ExcelHyperlinkType.DOCUMENT));

      assertEquals(
          "https://example.com/report",
          ExcelSheet.toPoiTarget(new ExcelHyperlink.Url("https://example.com/report")));
      assertEquals(
          "mailto:team@example.com",
          ExcelSheet.toPoiTarget(new ExcelHyperlink.Email("team@example.com")));
      assertEquals(
          Path.of("/tmp/report.xlsx").toUri().toASCIIString(),
          ExcelSheet.toPoiTarget(new ExcelHyperlink.File("/tmp/report.xlsx")));
      assertEquals(
          "support/budget%20backup.xlsx",
          ExcelSheet.toPoiTarget(new ExcelHyperlink.File("support/budget backup.xlsx")));
      assertEquals("Budget!B4", ExcelSheet.toPoiTarget(new ExcelHyperlink.Document("Budget!B4")));
    }
  }

  @Test
  void derivesFormulaHealthFindingsAcrossExternalVolatileErrorAndFailureCases() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.formula("INDIRECT(\"[External.xlsx]Sheet1!A1\")"));
      sheet.setCell("A2", ExcelCellValue.formula("NOW()"));
      sheet.setCell("A3", ExcelCellValue.formula("1/0"));

      List<WorkbookAnalysis.AnalysisFinding> findings = sheet.formulaHealthFindings();

      assertEquals(3, sheet.formulaCellCount());
      assertTrue(
          findings.stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .containsAll(
                  List.of(
                      WorkbookAnalysis.AnalysisFindingCode.FORMULA_EXTERNAL_REFERENCE,
                      WorkbookAnalysis.AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                      WorkbookAnalysis.AnalysisFindingCode.FORMULA_ERROR_RESULT)));
    }
  }

  @Test
  void derivesFormulaEvaluationFailureFindingWithFallbackExceptionMessage() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      poiSheet.createRow(0).createCell(0).setCellFormula("1+1");
      ExcelSheet sheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.alwaysFail(new IllegalStateException()));

      List<WorkbookAnalysis.AnalysisFinding> findings = sheet.formulaHealthFindings();

      assertEquals(1, findings.size());
      WorkbookAnalysis.AnalysisFinding finding = findings.getFirst();
      assertEquals(WorkbookAnalysis.AnalysisFindingCode.FORMULA_EVALUATION_FAILURE, finding.code());
      assertEquals(List.of("1+1", "IllegalStateException"), finding.evidence());
      assertEquals("Formula evaluation failed: IllegalStateException", finding.message());
    }
  }

  @Test
  void formulaHealthFindingsIgnoreSuccessfulAndNullEvaluations() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet successSheet = poiWorkbook.createSheet("Success");
      successSheet.createRow(0).createCell(0).setCellFormula("1+1");
      ExcelSheet successful =
          new ExcelSheet(
              successSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              poiWorkbook.getCreationHelper().createFormulaEvaluator());
      assertEquals(List.of(), successful.formulaHealthFindings());

      Sheet nullSheet = poiWorkbook.createSheet("NullEval");
      nullSheet.createRow(0).createCell(0).setCellFormula("1+1");
      FormulaEvaluator baseEvaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet nullEvaluated =
          new ExcelSheet(
              nullSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.nullEvaluation(baseEvaluator));
      assertEquals(List.of(), nullEvaluated.formulaHealthFindings());
    }
  }

  @Test
  void derivesHyperlinkHealthFindingsAcrossMalformedAndDocumentTargets() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      poiWorkbook.createSheet("Quarter 1");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      Path reachableFile = Files.createTempFile("gridgrind-hyperlink-health-", ".xlsx");
      Cell validFileCell = poiSheet.createRow(0).createCell(0);
      validFileCell.setHyperlink(
          hyperlink(poiWorkbook, HyperlinkType.FILE, reachableFile.toString()));
      Cell missingSheetCell = poiSheet.createRow(1).createCell(0);
      missingSheetCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.DOCUMENT, "Missing!A1"));
      Cell invalidDocumentCell = poiSheet.createRow(2).createCell(0);
      invalidDocumentCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.DOCUMENT, "Budget!"));
      Cell invalidDocumentRangeCell = poiSheet.createRow(3).createCell(0);
      invalidDocumentRangeCell.setHyperlink(
          hyperlink(poiWorkbook, HyperlinkType.DOCUMENT, "Budget!A1:"));
      Cell quotedDocumentCell = poiSheet.createRow(4).createCell(0);
      quotedDocumentCell.setHyperlink(
          hyperlink(poiWorkbook, HyperlinkType.DOCUMENT, "'Quarter 1'!A1"));

      List<WorkbookAnalysis.AnalysisFinding> findings;
      try {
        findings = sheet.hyperlinkHealthFindings();
      } finally {
        Files.deleteIfExists(reachableFile);
      }

      assertEquals(5, sheet.hyperlinkCount());
      assertTrue(
          findings.stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .containsAll(
                  List.of(
                      WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MISSING_DOCUMENT_SHEET,
                      WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET)));
      assertTrue(
          findings.stream()
              .noneMatch(
                  finding ->
                      finding.location() instanceof WorkbookAnalysis.AnalysisLocation.Cell cell
                          && "A5".equals(cell.address())));
    }
  }

  @Test
  void fileHyperlinkFindingsCoverMalformedMissingAndUnresolvedTargets() {
    WorkbookAnalysis.AnalysisLocation.Cell location =
        new WorkbookAnalysis.AnalysisLocation.Cell("Budget", "A1");
    WorkbookLocation storedWorkbook =
        new WorkbookLocation.StoredWorkbook(
            Path.of("tmp", "file-hyperlink-findings", "Budget.xlsx").toAbsolutePath());

    List<WorkbookAnalysis.AnalysisFinding> malformed =
        ExcelSheet.fileHyperlinkFindings(location, "https://example.com/report", storedWorkbook);
    assertEquals(1, malformed.size());
    assertEquals(
        WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET,
        malformed.getFirst().code());

    List<WorkbookAnalysis.AnalysisFinding> unresolved =
        ExcelSheet.fileHyperlinkFindings(
            location, "reports/q1.xlsx", new WorkbookLocation.UnsavedWorkbook());
    assertEquals(1, unresolved.size());
    assertEquals(
        WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET,
        unresolved.getFirst().code());

    List<WorkbookAnalysis.AnalysisFinding> missing =
        ExcelSheet.fileHyperlinkFindings(location, "reports/q1.xlsx", storedWorkbook);
    assertEquals(1, missing.size());
    assertEquals(
        WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MISSING_FILE_TARGET,
        missing.getFirst().code());
  }

  @Test
  void externalHyperlinkFindingsCoverMalformedUrlAndEmailTargets() {
    WorkbookAnalysis.AnalysisLocation.Cell location =
        new WorkbookAnalysis.AnalysisLocation.Cell("Budget", "A1");

    List<WorkbookAnalysis.AnalysisFinding> malformedUrl =
        ExcelSheet.externalHyperlinkFindings(location, "example.com/report", "URL", false);
    assertEquals(1, malformedUrl.size());
    assertEquals(
        WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET,
        malformedUrl.getFirst().code());

    List<WorkbookAnalysis.AnalysisFinding> malformedEmail =
        ExcelSheet.externalHyperlinkFindings(location, "mailto:", "EMAIL", false);
    assertEquals(1, malformedEmail.size());
    assertEquals(
        WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET,
        malformedEmail.getFirst().code());
  }

  @Test
  void normalizesSparseWindowsAndInvalidHyperlinkHelpers() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Sparse");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.text("first"));
      sheet.setCell("A3", ExcelCellValue.text("third"));

      WorkbookReadResult.Window window = sheet.window("A1", 3, 1);
      assertEquals("A2", window.rows().get(1).cells().getFirst().address());
      assertEquals("BLANK", window.rows().get(1).cells().getFirst().effectiveType());

      WorkbookReadResult.SheetLayout layout = sheet.layout();
      assertEquals(3, layout.rows().size());
      assertEquals(poiSheet.getDefaultRowHeightInPoints(), layout.rows().get(1).heightPoints());

      assertNull(ExcelSheet.hyperlink(HyperlinkType.URL, "example.com/report"));
    }
  }

  @Test
  void helperMethodsHandleFormulaAndHyperlinkEdgeCases() throws Exception {
    assertTrue(ExcelSheet.containsExternalWorkbookReference("[Book.xlsx]Sheet1!A1"));
    assertFalse(ExcelSheet.containsExternalWorkbookReference("SUM(A1:A2)"));
    assertFalse(ExcelSheet.containsExternalWorkbookReference("[Book.xlsx"));
    assertFalse(ExcelSheet.containsExternalWorkbookReference("Book.xlsx]"));

    assertEquals(
        List.of("NOW", "INDIRECT"), ExcelSheet.volatileFunctions("NOW()+INDIRECT(\"A1\")"));
    assertEquals(List.of(), ExcelSheet.volatileFunctions("SUM(A1:A2)"));

    assertEquals("Quarter 1", ExcelSheet.unquoteSheetName("'Quarter 1'"));
    assertEquals("Budget", ExcelSheet.unquoteSheetName("Budget"));
    assertEquals("'", ExcelSheet.unquoteSheetName("'"));
    assertEquals("'Budget", ExcelSheet.unquoteSheetName("'Budget"));
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      ExcelSheet sheet =
          new ExcelSheet(
              poiWorkbook.createSheet("Helpers"),
              new WorkbookStyleRegistry(poiWorkbook),
              poiWorkbook.getCreationHelper().createFormulaEvaluator());
      assertEquals("IllegalStateException", sheet.exceptionMessage(new IllegalStateException()));
      assertEquals(
          "display failure", sheet.exceptionMessage(new IllegalStateException("display failure")));
      WorkbookAnalysis.AnalysisLocation.Cell location =
          new WorkbookAnalysis.AnalysisLocation.Cell("Helpers", "A1");
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET,
          sheet
              .hyperlinkTargetFindings(
                  location, HyperlinkType.URL, " ", new WorkbookLocation.UnsavedWorkbook())
              .getFirst()
              .code());
      assertEquals(
          List.of(),
          sheet.hyperlinkTargetFindings(
              location,
              HyperlinkType.EMAIL,
              "team@example.com",
              new WorkbookLocation.UnsavedWorkbook()));
      assertEquals(
          List.of(),
          sheet.hyperlinkTargetFindings(
              location, HyperlinkType.NONE, "ignored", new WorkbookLocation.UnsavedWorkbook()));
    }

    assertNull(ExcelSheet.hyperlink((Cell) null));
    assertFalse(ExcelSheet.hasUsableHyperlink((org.apache.poi.ss.usermodel.Hyperlink) null));
    assertFalse(ExcelSheet.hasUsableHyperlinkType(null));
    assertFalse(ExcelSheet.hasUsableHyperlinkType(HyperlinkType.NONE));
    assertTrue(ExcelSheet.hasUsableHyperlinkType(HyperlinkType.URL));
    assertNull(ExcelSheet.hyperlink(hyperlinkWithNullType()));

    assertTrue(ExcelSheet.hasMissingHyperlinkTarget(null));
    assertTrue(ExcelSheet.hasMissingHyperlinkTarget(" "));
    assertFalse(ExcelSheet.hasMissingHyperlinkTarget("Budget!A1"));
    assertNull(ExcelSheet.comment((Cell) null));
  }

  @Test
  void validateDocumentHyperlinkTargetHandlesMalformedMissingAndRangeTargets() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      poiWorkbook.createSheet("Quarter 1");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);
      WorkbookAnalysis.AnalysisLocation.Cell location =
          new WorkbookAnalysis.AnalysisLocation.Cell("Budget", "A1");

      List<WorkbookAnalysis.AnalysisFinding> invalidStructure = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "!A1", invalidStructure);
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET,
          invalidStructure.getFirst().code());

      List<WorkbookAnalysis.AnalysisFinding> missingSheet = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "Missing!A1", missingSheet);
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MISSING_DOCUMENT_SHEET,
          missingSheet.getFirst().code());

      List<WorkbookAnalysis.AnalysisFinding> invalidRange = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "Budget!A1:", invalidRange);
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET,
          invalidRange.getFirst().code());

      List<WorkbookAnalysis.AnalysisFinding> validRange = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "Budget!A1:B2", validRange);
      assertEquals(List.of(), validRange);

      List<WorkbookAnalysis.AnalysisFinding> quotedValid = new ArrayList<>();
      sheet.validateDocumentHyperlinkTarget(location, "'Quarter 1'!A1", quotedValid);
      assertEquals(List.of(), quotedValid);
    }
  }

  @Test
  void hyperlinkHealthTreatsEmptyCellsAndNoneLinksAsNonFindings() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      poiSheet.createRow(0).createCell(0).setCellValue("plain");
      Cell noneCell = poiSheet.createRow(1).createCell(0);
      noneCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.NONE, "ignored"));
      Cell validUrlCell = poiSheet.createRow(2).createCell(0);
      validUrlCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.URL, "https://example.com"));
      Cell validEmailCell = poiSheet.createRow(3).createCell(0);
      validEmailCell.setHyperlink(hyperlink(poiWorkbook, HyperlinkType.EMAIL, "team@example.com"));

      assertEquals(2, sheet.hyperlinkCount());
      assertEquals(List.of(), sheet.hyperlinkHealthFindings());
    }
  }

  private org.apache.poi.ss.usermodel.Hyperlink hyperlink(
      XSSFWorkbook workbook, HyperlinkType hyperlinkType, String address) {
    org.apache.poi.ss.usermodel.Hyperlink hyperlink =
        workbook.getCreationHelper().createHyperlink(hyperlinkType);
    if (hyperlinkType != HyperlinkType.NONE) {
      hyperlink.setAddress(address);
    }
    return hyperlink;
  }

  private static org.apache.poi.ss.usermodel.Hyperlink hyperlinkWithNullType() {
    return new org.apache.poi.ss.usermodel.Hyperlink() {
      @Override
      public HyperlinkType getType() {
        return null;
      }

      @Override
      public String getAddress() {
        return "ignored";
      }

      @Override
      public void setAddress(String address) {}

      @Override
      public String getLabel() {
        return null;
      }

      @Override
      public void setLabel(String label) {}

      @Override
      public int getFirstRow() {
        return 0;
      }

      @Override
      public void setFirstRow(int row) {}

      @Override
      public int getLastRow() {
        return 0;
      }

      @Override
      public void setLastRow(int row) {}

      @Override
      public int getFirstColumn() {
        return 0;
      }

      @Override
      public void setFirstColumn(int column) {}

      @Override
      public int getLastColumn() {
        return 0;
      }

      @Override
      public void setLastColumn(int column) {}
    };
  }

  private Comment comment(
      XSSFWorkbook workbook, Sheet sheet, String text, String author, boolean visible) {
    Comment comment = emptyComment(workbook, sheet);
    comment.setString(workbook.getCreationHelper().createRichTextString(text));
    comment.setAuthor(author);
    comment.setVisible(visible);
    return comment;
  }

  private Comment emptyComment(XSSFWorkbook workbook, Sheet sheet) {
    ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
    anchor.setRow1(0);
    anchor.setRow2(3);
    anchor.setCol1(0);
    anchor.setCol2(3);
    return sheet.createDrawingPatriarch().createCellComment(anchor);
  }

  @Test
  void snapshotCellReturnsBlankForUnwrittenCells() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.text("Header"));

      // A cell in a non-existent row returns blank.
      ExcelCellSnapshot blankInNewRow = sheet.snapshotCell("C5");
      assertInstanceOf(ExcelCellSnapshot.BlankSnapshot.class, blankInNewRow);
      assertEquals("C5", blankInNewRow.address());

      // A cell in a row that exists but in a column that has not been written returns blank.
      ExcelCellSnapshot blankInExistingRow = sheet.snapshotCell("B1");
      assertInstanceOf(ExcelCellSnapshot.BlankSnapshot.class, blankInExistingRow);
      assertEquals("B1", blankInExistingRow.address());
    }
  }
}
