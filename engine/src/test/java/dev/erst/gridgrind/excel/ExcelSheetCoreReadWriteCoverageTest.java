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
import java.util.Optional;
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
      ExcelSheetCells cells = sheet.cells();

      assertSame(
          cells,
          cells.appendRow(
              ExcelCellValue.text("Name"),
              ExcelCellValue.number(42.5),
              ExcelCellValue.bool(true),
              ExcelCellValue.formula("B1*2"),
              ExcelCellValue.formula("TRUE()"),
              ExcelCellValue.formula("\"Hi\"")));
      sheet
          .cells()
          .appendRow(
              ExcelCellValue.blank(),
              ExcelCellValue.date(LocalDate.of(2026, 3, 23)),
              ExcelCellValue.dateTime(LocalDateTime.of(2026, 3, 23, 14, 15, 16)),
              ExcelCellValue.formula("1/0"));
      sheet.cells().setCell("A3", ExcelCellValue.formula("B3+1"));
      sheet.cells().setCell("B3", ExcelCellValue.date(LocalDate.of(2026, 3, 24)));
      sheet.cells().setCell("C3", ExcelCellValue.dateTime(LocalDateTime.of(2026, 3, 24, 9, 0)));
      sheet.cells().setCell("D3", ExcelCellValue.formula("D3+1"));
      sheet.columns().autoSize();

      Row errorRow = poiSheet.createRow(3);
      Cell errorCell = errorRow.createCell(0);
      errorCell.setCellErrorValue(FormulaError.DIV0.getCode());

      assertEquals("Budget", sheet.name());
      assertEquals("Name", sheet.cells().text("A1"));
      assertEquals(85.0, sheet.cells().number("D1"));
      assertTrue(sheet.cells().bool("C1"));
      assertTrue(sheet.cells().bool("E1"));
      assertEquals("B1*2", sheet.cells().formula("D1"));
      assertEquals(4, sheet.rows().physicalCount());
      assertEquals(3, sheet.rows().lastIndex());
      assertEquals(5, sheet.columns().lastIndex());

      ExcelCellSnapshot.TextSnapshot textSnapshot =
          (ExcelCellSnapshot.TextSnapshot) sheet.cells().snapshotCell("A1");
      assertEquals("TEXT", textSnapshot.type());
      assertEquals("Name", textSnapshot.textValue());
      assertNull(textSnapshot.richText());

      ExcelCellSnapshot.NumberSnapshot numberSnapshot =
          (ExcelCellSnapshot.NumberSnapshot) sheet.cells().snapshotCell("B1");
      assertEquals("NUMBER", numberSnapshot.type());
      assertEquals(42.5, numberSnapshot.numberValue());

      ExcelCellSnapshot.BooleanSnapshot booleanSnapshot =
          (ExcelCellSnapshot.BooleanSnapshot) sheet.cells().snapshotCell("C1");
      assertEquals("BOOLEAN", booleanSnapshot.type());
      assertTrue(booleanSnapshot.booleanValue());

      ExcelCellSnapshot.BlankSnapshot blankSnapshot =
          (ExcelCellSnapshot.BlankSnapshot) sheet.cells().snapshotCell("A2");
      assertEquals("BLANK", blankSnapshot.type());

      ExcelCellSnapshot.FormulaSnapshot stringFormulaSnapshot =
          (ExcelCellSnapshot.FormulaSnapshot) sheet.cells().snapshotCell("F1");
      assertEquals("FORMULA", stringFormulaSnapshot.type());
      assertEquals("\"Hi\"", stringFormulaSnapshot.formula());
      assertEquals(
          "Hi", ((ExcelCellSnapshot.TextSnapshot) stringFormulaSnapshot.evaluation()).textValue());

      ExcelCellSnapshot.FormulaSnapshot errorFormulaSnapshot =
          (ExcelCellSnapshot.FormulaSnapshot) sheet.cells().snapshotCell("D2");
      assertEquals("FORMULA", errorFormulaSnapshot.type());
      assertEquals(
          "#DIV/0!",
          ((ExcelCellSnapshot.ErrorSnapshot) errorFormulaSnapshot.evaluation()).errorValue());

      ExcelCellSnapshot.ErrorSnapshot errorSnapshot =
          (ExcelCellSnapshot.ErrorSnapshot) sheet.cells().snapshotCell("A4");
      assertEquals("ERROR", errorSnapshot.type());
      assertEquals("#DIV/0!", errorSnapshot.errorValue());
      ExcelCellSnapshot.FormulaSnapshot circularFormulaSnapshot =
          (ExcelCellSnapshot.FormulaSnapshot) sheet.cells().snapshotCell("D3");
      assertEquals(
          "#CIRCULAR_REF!",
          ((ExcelCellSnapshot.ErrorSnapshot) circularFormulaSnapshot.evaluation()).errorValue());

      List<ExcelPreviewRow> preview = sheet.cells().preview(4, 6);
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

      assertThrows(
          NullPointerException.class, () -> sheet.cells().setCell(null, ExcelCellValue.text("x")));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.cells().setCell(" ", ExcelCellValue.text("x")));
      assertThrows(NullPointerException.class, () -> sheet.cells().setCell("A1", null));
      assertThrows(
          NullPointerException.class,
          () -> sheet.cells().setRange(null, List.of(List.of(ExcelCellValue.text("x")))));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.cells().setRange(" ", List.of(List.of(ExcelCellValue.text("x")))));
      assertThrows(NullPointerException.class, () -> sheet.cells().setRange("A1", null));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().setRange("A1", List.of()));
      assertThrows(NullPointerException.class, () -> sheet.cells().clearRange(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().clearRange(" "));
      assertThrows(NullPointerException.class, () -> sheet.layout().mergeCells(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().mergeCells(" "));
      assertThrows(NullPointerException.class, () -> sheet.layout().unmergeCells(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().unmergeCells(" "));
      assertThrows(IllegalArgumentException.class, () -> sheet.columns().setWidth(-1, 0, 16.0));
      assertThrows(IllegalArgumentException.class, () -> sheet.columns().setWidth(1, 0, 16.0));
      assertThrows(IllegalArgumentException.class, () -> sheet.columns().setWidth(0, 0, 0.0));
      assertThrows(IllegalArgumentException.class, () -> sheet.columns().setWidth(0, 0, 256.0));
      assertThrows(
          IllegalArgumentException.class, () -> sheet.columns().setWidth(0, 0, Double.MIN_VALUE));
      assertThrows(
          IllegalArgumentException.class, () -> sheet.columns().setWidth(0, 0, Double.NaN));
      assertThrows(IllegalArgumentException.class, () -> sheet.rows().setHeight(-1, 0, 28.5));
      assertThrows(IllegalArgumentException.class, () -> sheet.rows().setHeight(1, 0, 28.5));
      assertThrows(IllegalArgumentException.class, () -> sheet.rows().setHeight(0, 0, 0.0));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .rows()
                  .setHeight(0, 0, Math.nextUp(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS)));
      assertThrows(
          IllegalArgumentException.class, () -> sheet.rows().setHeight(0, 0, Double.MIN_VALUE));
      assertThrows(IllegalArgumentException.class, () -> sheet.rows().setHeight(0, 0, Double.NaN));
      assertThrows(NullPointerException.class, () -> sheet.layout().setPane(null));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new ExcelSheetPane.Frozen(-1, 0, 0, 0)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new ExcelSheetPane.Frozen(0, 0, 0, 0)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new ExcelSheetPane.Frozen(0, 1, 1, 1)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new ExcelSheetPane.Frozen(1, 0, 1, 1)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new ExcelSheetPane.Frozen(2, 1, 1, 1)));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.layout().setPane(new ExcelSheetPane.Frozen(1, 2, 1, 1)));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().setZoom(9));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().setZoom(401));
      assertThrows(NullPointerException.class, () -> sheet.layout().setPrintLayout(null));
      assertThrows(
          NullPointerException.class,
          () -> sheet.cells().applyStyle(null, ExcelCellStyle.numberFormat("0")));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.cells().applyStyle(" ", ExcelCellStyle.numberFormat("0")));
      assertThrows(NullPointerException.class, () -> sheet.cells().applyStyle("A1", null));
      assertThrows(
          NullPointerException.class,
          () ->
              sheet
                  .annotations()
                  .setHyperlink(null, new ExcelHyperlink.Url("https://example.com")));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet.annotations().setHyperlink(" ", new ExcelHyperlink.Url("https://example.com")));
      assertThrows(NullPointerException.class, () -> sheet.annotations().setHyperlink("A1", null));
      assertThrows(NullPointerException.class, () -> sheet.annotations().clearHyperlink(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.annotations().clearHyperlink(" "));
      assertThrows(
          NullPointerException.class,
          () ->
              sheet.annotations().setComment(null, new ExcelComment("Review", "GridGrind", false)));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet.annotations().setComment(" ", new ExcelComment("Review", "GridGrind", false)));
      assertThrows(NullPointerException.class, () -> sheet.annotations().setComment("A1", null));
      assertThrows(NullPointerException.class, () -> sheet.annotations().clearComment(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.annotations().clearComment(" "));
      assertThrows(
          NullPointerException.class, () -> sheet.cells().appendRow((ExcelCellValue[]) null));
      assertThrows(
          NullPointerException.class,
          () -> sheet.cells().appendRow(ExcelCellValue.text("x"), null));
    }
  }

  @Test
  void validatesAddressAndRangeParsing() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      assertThrows(NullPointerException.class, () -> sheet.cells().snapshotCell(null));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().snapshotCell(" "));
      InvalidCellAddressException invalidSetCell =
          assertThrows(
              InvalidCellAddressException.class,
              () -> sheet.cells().setCell(":", ExcelCellValue.text("x")));
      assertEquals(":", invalidSetCell.address());
      InvalidCellAddressException invalidSnapshotCell =
          assertThrows(InvalidCellAddressException.class, () -> sheet.cells().snapshotCell(":"));
      assertEquals(":", invalidSnapshotCell.address());
      InvalidCellAddressException badAddrSnapshot =
          assertThrows(
              InvalidCellAddressException.class, () -> sheet.cells().snapshotCell("BADADDR"));
      assertEquals("BADADDR", badAddrSnapshot.address());
      InvalidCellAddressException a0Snapshot =
          assertThrows(InvalidCellAddressException.class, () -> sheet.cells().snapshotCell("A0"));
      assertEquals("A0", a0Snapshot.address());
      InvalidCellAddressException numericOnlySnapshot =
          assertThrows(InvalidCellAddressException.class, () -> sheet.cells().snapshotCell("1"));
      assertEquals("1", numericOnlySnapshot.address());
      InvalidCellAddressException outOfBoundsRow =
          assertThrows(
              InvalidCellAddressException.class, () -> sheet.cells().snapshotCell("A1048577"));
      assertEquals("A1048577", outOfBoundsRow.address());
      InvalidCellAddressException outOfBoundsCol =
          assertThrows(InvalidCellAddressException.class, () -> sheet.cells().snapshotCell("XFE1"));
      assertEquals("XFE1", outOfBoundsCol.address());
      InvalidRangeAddressException invalidRangeSet =
          assertThrows(
              InvalidRangeAddressException.class,
              () -> sheet.cells().setRange("A1:", List.of(List.of(ExcelCellValue.text("x")))));
      assertEquals("A1:", invalidRangeSet.range());
      InvalidRangeAddressException invalidRangeClear =
          assertThrows(
              InvalidRangeAddressException.class, () -> sheet.cells().clearRange("A1:B2:C3"));
      assertEquals("A1:B2:C3", invalidRangeClear.range());
      InvalidRangeAddressException invalidRangeMerge =
          assertThrows(InvalidRangeAddressException.class, () -> sheet.layout().mergeCells("A1:"));
      assertEquals("A1:", invalidRangeMerge.range());
      InvalidRangeAddressException invalidRangeUnmerge =
          assertThrows(
              InvalidRangeAddressException.class, () -> sheet.layout().unmergeCells("A1:"));
      assertEquals("A1:", invalidRangeUnmerge.range());
      assertThrows(
          InvalidRangeAddressException.class,
          () -> sheet.cells().applyStyle("A1:", ExcelCellStyle.numberFormat("0")));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().mergeCells("A1"));
      assertThrows(IllegalArgumentException.class, () -> sheet.layout().unmergeCells("A1:B2"));
    }
  }

  @Test
  void groupsAndUngroupsRowsAndColumnsThroughSheetFacade() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);
      ExcelSheetRows rows = sheet.rows();
      ExcelSheetColumns columns = sheet.columns();

      sheet.cells().setCell("A1", ExcelCellValue.text("Header"));

      assertSame(rows, rows.group(new ExcelRowSpan(1, 3), true));
      assertSame(columns, columns.group(new ExcelColumnSpan(1, 3), true));
      assertSame(rows, rows.ungroup(new ExcelRowSpan(1, 3)));
      assertSame(columns, columns.ungroup(new ExcelColumnSpan(1, 3)));

      WorkbookSheetResult.SheetLayout layout = sheet.layout().snapshot();
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

      assertThrows(IllegalArgumentException.class, () -> sheet.cells().preview(0, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().preview(1, 0));
      assertEquals(List.of(), sheet.cells().preview(3, 3));
      CellNotFoundException missingCell =
          assertThrows(CellNotFoundException.class, () -> sheet.cells().text("A1"));
      assertEquals("A1", missingCell.address());
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().text(" "));

      sheet.cells().setCell("A1", ExcelCellValue.text("Name"));
      sheet.cells().setCell("B1", ExcelCellValue.formula("TRUE()"));
      sheet.cells().setCell("C1", ExcelCellValue.formula("1+1"));

      assertThrows(CellNotFoundException.class, () -> sheet.cells().text("B2"));
      assertThrows(CellNotFoundException.class, () -> sheet.cells().text("D1"));
      assertThrows(IllegalStateException.class, () -> sheet.cells().number("A1"));
      assertThrows(IllegalStateException.class, () -> sheet.cells().number("B1"));
      assertThrows(IllegalStateException.class, () -> sheet.cells().bool("C1"));
      assertThrows(IllegalStateException.class, () -> sheet.cells().formula("A1"));
    }
  }

  @Test
  void supportsRangeWritesClearsStylesAndStyleAwarePreview() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);
      ExcelSheetCells cells = sheet.cells();

      assertSame(
          cells,
          cells.setRange(
              "B2:A1",
              List.of(
                  List.of(ExcelCellValue.text("Item"), ExcelCellValue.number(42.0)),
                  List.of(ExcelCellValue.text("Tax"), ExcelCellValue.number(8.0)))));
      sheet
          .cells()
          .applyStyle(
              "A1:B1",
              new ExcelCellStyle(
                  Optional.of("#,##0.00"),
                  Optional.of(
                      new ExcelCellAlignment(
                          Optional.of(true),
                          Optional.of(ExcelHorizontalAlignment.CENTER),
                          Optional.of(ExcelVerticalAlignment.TOP),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.of(
                      new ExcelCellFont(
                          Optional.of(true),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty(),
                  Optional.empty(),
                  Optional.empty()));
      sheet.cells().applyStyle("C1", ExcelCellStyle.emphasis(null, true));

      ExcelCellSnapshot styledValue = sheet.cells().snapshotCell("A1");
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

      List<ExcelPreviewRow> preview = sheet.cells().preview(2, 3);
      assertTrue(preview.getFirst().cells().stream().anyMatch(cell -> "C1".equals(cell.address())));
      assertEquals("BLANK", sheet.cells().snapshotCell("C1").type());
      assertTrue(sheet.cells().snapshotCell("C1").style().font().italic());

      sheet.cells().clearRange("A2:B2");

      ExcelCellSnapshot cleared = sheet.cells().snapshotCell("A2");
      assertEquals("BLANK", cleared.type());
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
          () -> sheet.cells().setRange("A1:B2", List.of(List.of(ExcelCellValue.text("x")))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .cells()
                  .setRange("A1:B2", List.of(List.of(), List.of(ExcelCellValue.text("x")))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .cells()
                  .setRange(
                      "A1:B2",
                      List.of(
                          List.of(ExcelCellValue.text("x")),
                          List.of(ExcelCellValue.text("y"), ExcelCellValue.text("z")))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .cells()
                  .setRange(
                      "A1:B2",
                      List.of(
                          List.of(ExcelCellValue.text("x")), List.of(ExcelCellValue.text("y")))));
      List<List<ExcelCellValue>> rowsWithNull = new ArrayList<>();
      List<ExcelCellValue> rowWithNull = new ArrayList<>();
      rowWithNull.add(null);
      rowsWithNull.add(rowWithNull);
      assertThrows(NullPointerException.class, () -> sheet.cells().setRange("A1", rowsWithNull));
    }
  }

  @Test
  void mergesFormattingDepthStylesAndPreservesExistingAttributes() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.cells().setCell("A1", ExcelCellValue.text("Item"));
      sheet
          .cells()
          .applyStyle(
              "A1",
              new ExcelCellStyle(
                  Optional.empty(),
                  Optional.of(
                      new ExcelCellAlignment(
                          Optional.of(true),
                          Optional.of(ExcelHorizontalAlignment.CENTER),
                          Optional.of(ExcelVerticalAlignment.TOP),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.of(
                      new ExcelCellFont(
                          Optional.of(true),
                          Optional.of(false),
                          Optional.of("Aptos"),
                          Optional.of(new ExcelFontHeight(280)),
                          Optional.of(ExcelColor.rgb("#1F4E78")),
                          Optional.of(true),
                          Optional.of(false))),
                  Optional.of(
                      ExcelCellFill.patternForeground(
                          ExcelFillPattern.SOLID, ExcelColor.rgb("#FFF2CC"))),
                  Optional.of(
                      new ExcelBorder(
                          Optional.ofNullable(new ExcelBorderSide(ExcelBorderStyle.THIN)),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty()));
      sheet
          .cells()
          .applyStyle(
              "A1",
              new ExcelCellStyle(
                  Optional.empty(),
                  Optional.empty(),
                  Optional.of(
                      new ExcelCellFont(
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.of(true))),
                  Optional.empty(),
                  Optional.of(
                      new ExcelBorder(
                          Optional.empty(),
                          Optional.empty(),
                          Optional.ofNullable(new ExcelBorderSide(ExcelBorderStyle.DOUBLE)),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty()));

      ExcelCellSnapshot styled = sheet.cells().snapshotCell("A1");
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

      sheet.annotations().setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
      sheet.annotations().setComment("A1", new ExcelComment("Review", "GridGrind", true));

      ExcelCellSnapshot.BlankSnapshot snapshot =
          (ExcelCellSnapshot.BlankSnapshot) sheet.cells().snapshotCell("A1");
      assertEquals(
          new ExcelHyperlink.Url("https://example.com/report"),
          snapshot.metadata().hyperlink().orElseThrow());
      assertEquals(
          new ExcelComment("Review", "GridGrind", true),
          snapshot.metadata().comment().orElseThrow().toPlainComment());

      List<ExcelPreviewRow> preview = sheet.cells().preview(1, 1);
      assertEquals(1, preview.size());
      assertEquals("A1", preview.getFirst().cells().getFirst().address());

      sheet.annotations().clearHyperlink("A1");
      sheet.annotations().clearComment("A1");
      ExcelCellSnapshot.BlankSnapshot clearedMetadata =
          (ExcelCellSnapshot.BlankSnapshot) sheet.cells().snapshotCell("A1");
      assertTrue(clearedMetadata.metadata().hyperlink().isEmpty());
      assertTrue(clearedMetadata.metadata().comment().isEmpty());

      sheet.annotations().setHyperlink("A1", new ExcelHyperlink.Document("Budget!B4"));
      sheet.annotations().setComment("A1", new ExcelComment("Again", "GridGrind", false));
      sheet.cells().clearRange("A1");
      ExcelCellSnapshot.BlankSnapshot clearedRange =
          (ExcelCellSnapshot.BlankSnapshot) sheet.cells().snapshotCell("A1");
      assertTrue(clearedRange.metadata().hyperlink().isEmpty());
      assertTrue(clearedRange.metadata().comment().isEmpty());

      // clearHyperlink and clearComment are no-ops on cells that do not physically exist
      assertDoesNotThrow(() -> sheet.annotations().clearHyperlink("B2"));
      assertDoesNotThrow(() -> sheet.annotations().clearComment("B2"));
      // calling again on B2 (still non-existent) must still be a no-op, not throw
      assertDoesNotThrow(() -> sheet.annotations().clearHyperlink("B2"));
      assertDoesNotThrow(() -> sheet.annotations().clearComment("B2"));
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
          () ->
              sheet
                  .annotations()
                  .setHyperlink("A1", new ExcelHyperlink.File("support/budget backup.xlsx")));

      Cell cell = poiSheet.getRow(0).getCell(0);
      assertEquals("support/budget%20backup.xlsx", cell.getHyperlink().getAddress());
      assertEquals(
          Optional.of(new ExcelHyperlink.File("support/budget backup.xlsx")),
          ExcelSheetAnnotationSupport.hyperlink(cell));
    }
  }

  @Test
  void clearRangeIsNoOpOnNeverWrittenCells() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Data");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.cells().setCell("A1", ExcelCellValue.text("anchor"));

      int rowsBefore = sheet.rows().physicalCount();
      int lastRowBefore = sheet.rows().lastIndex();
      int lastColBefore = sheet.columns().lastIndex();

      // clear a range that has never been written
      sheet.cells().clearRange("B2:E5");

      // physicalRowCount, lastRowIndex, and lastColumnIndex must not change
      assertEquals(rowsBefore, sheet.rows().physicalCount(), "physicalRowCount must not increase");
      assertEquals(lastRowBefore, sheet.rows().lastIndex(), "lastRowIndex must not change");
      assertEquals(lastColBefore, sheet.columns().lastIndex(), "lastColumnIndex must not change");
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
      sheet.cells().setCell("B2", ExcelCellValue.text("present"));
      assertDoesNotThrow(() -> sheet.cells().clearRange("B2:C2"));
      // B2 must now be blank after the clear
      ExcelCellSnapshot b2 = sheet.cells().snapshotCell("B2");
      assertEquals("BLANK", b2.type());
    }
  }

  @Test
  void introspectsWindowsSelectionsAndLayoutAcrossSparseMetadata() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      WorkbookSheetResult.SheetLayout emptyLayout = sheet.layout().snapshot();
      assertEquals(new ExcelSheetPane.None(), emptyLayout.pane());
      assertEquals(100, emptyLayout.zoomPercent());
      assertEquals(ExcelSheetDisplay.defaults(), emptyLayout.presentation().display());
      assertEquals(
          ExcelSheetOutlineSummary.defaults(), emptyLayout.presentation().outlineSummary());
      assertEquals(ExcelSheetDefaults.defaults(), emptyLayout.presentation().sheetDefaults());
      assertEquals(List.of(), emptyLayout.columns());
      assertEquals(List.of(), emptyLayout.rows());

      sheet.cells().setCell("B2", ExcelCellValue.text("Center"));
      sheet.annotations().setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
      sheet.annotations().setComment("C3", new ExcelComment("Review", "GridGrind", false));
      sheet.columns().setWidth(0, 0, 12.5);
      sheet.rows().setHeight(0, 0, 19.5);
      sheet
          .layout()
          .setPresentation(
              new ExcelSheetPresentation(
                  new ExcelSheetDisplay(false, false, false, true, true),
                  Optional.of(ExcelColor.rgb("#112233")),
                  new ExcelSheetOutlineSummary(false, false),
                  new ExcelSheetDefaults(11, 18.5d),
                  List.of(
                      new ExcelIgnoredError(
                          "A1:B2", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)))));

      WorkbookSheetResult.Window window = sheet.cells().window("A1", 3, 3);
      assertEquals("A1", window.rows().getFirst().cells().getFirst().address());
      assertEquals("B2", window.rows().get(1).cells().get(1).address());
      assertEquals("Center", window.rows().get(1).cells().get(1).displayValue());
      assertEquals("C3", window.rows().get(2).cells().get(2).address());
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().window("A1", 0, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().window("A1", 1, 0));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().window("A1", 1048577, 1));
      assertThrows(IllegalArgumentException.class, () -> sheet.cells().window("A1", 1, 16385));

      List<WorkbookSheetResult.CellHyperlink> allHyperlinks =
          sheet.annotations().hyperlinks(new ExcelCellSelection.AllUsedCells());
      assertEquals(1, allHyperlinks.size());
      assertEquals("A1", allHyperlinks.getFirst().address());

      List<WorkbookSheetResult.CellHyperlink> selectedHyperlinks =
          sheet
              .annotations()
              .hyperlinks(new ExcelCellSelection.Selected(List.of("A1", "B2", "B9", "D3")));
      assertEquals(1, selectedHyperlinks.size());
      assertEquals("A1", selectedHyperlinks.getFirst().address());

      List<WorkbookSheetResult.CellComment> selectedComments =
          sheet
              .annotations()
              .comments(new ExcelCellSelection.Selected(List.of("C3", "B2", "A9", "D3")));
      assertEquals(1, selectedComments.size());
      assertEquals("C3", selectedComments.getFirst().address());

      WorkbookSheetResult.SheetLayout unfrozenLayout = sheet.layout().snapshot();
      assertEquals(new ExcelSheetPane.None(), unfrozenLayout.pane());
      assertEquals(100, unfrozenLayout.zoomPercent());
      assertEquals(
          new ExcelSheetDisplay(false, false, false, true, true),
          unfrozenLayout.presentation().display());
      assertEquals(
          Optional.of(ExcelColorSnapshot.rgb("#112233")), unfrozenLayout.presentation().tabColor());
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

      sheet.layout().setPane(new ExcelSheetPane.Frozen(1, 1, 1, 1));
      sheet.layout().setZoom(135);
      WorkbookSheetResult.SheetLayout frozenLayout = sheet.layout().snapshot();
      assertEquals(new ExcelSheetPane.Frozen(1, 1, 1, 1), frozenLayout.pane());
      assertEquals(135, frozenLayout.zoomPercent());

      poiSheet.createSplitPane(2000, 2000, 0, 0, PaneType.LOWER_RIGHT);
      WorkbookSheetResult.SheetLayout splitLayout = sheet.layout().snapshot();
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

      sheet.annotations().setHyperlink("F18", new ExcelHyperlink.Email("Report_Value@example.com"));
      sheet
          .annotations()
          .setHyperlink("F18", new ExcelHyperlink.Email("Summary.Total@example.com"));

      ExcelCellSnapshot.BlankSnapshot snapshot =
          (ExcelCellSnapshot.BlankSnapshot) sheet.cells().snapshotCell("F18");
      assertEquals(
          new ExcelHyperlink.Email("Summary.Total@example.com"),
          snapshot.metadata().hyperlink().orElseThrow());
    }
  }
}
