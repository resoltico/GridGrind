package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** ExcelSheet formatting, rich text, and append-row coverage. */
class ExcelSheetFormattingCoverageTest extends ExcelSheetTestSupport {
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
          sheet.snapshotCell("B5").metadata().comment().orElseThrow().toPlainComment());
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
              ExcelCellFill.patternForeground(ExcelFillPattern.SOLID, ExcelColor.rgb("#AABBCC")),
              new ExcelBorder(new ExcelBorderSide(ExcelBorderStyle.THIN), null, null, null, null),
              null));
      sheet.setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
      sheet.setComment("A1", new ExcelComment("Review", "GridGrind", false));

      ExcelCellSnapshot.TextSnapshot textSnapshot =
          (ExcelCellSnapshot.TextSnapshot)
              sheet.setCell("A1", ExcelCellValue.text("Quarterly report")).snapshotCell("A1");
      assertEquals("Quarterly report", textSnapshot.stringValue());
      assertEquals(rgb("#AABBCC"), fillForegroundColor(textSnapshot.style().fill()));
      assertEquals(ExcelBorderStyle.THIN, textSnapshot.style().border().top().style());
      assertEquals(
          ExcelHorizontalAlignment.CENTER, textSnapshot.style().alignment().horizontalAlignment());
      assertEquals(
          new ExcelHyperlink.Url("https://example.com/report"),
          textSnapshot.metadata().hyperlink().orElseThrow());
      assertEquals(
          new ExcelComment("Review", "GridGrind", false),
          textSnapshot.metadata().comment().orElseThrow().toPlainComment());

      ExcelCellSnapshot.BlankSnapshot blankSnapshot =
          (ExcelCellSnapshot.BlankSnapshot)
              sheet.setCell("A1", ExcelCellValue.blank()).snapshotCell("A1");
      assertEquals(rgb("#AABBCC"), fillForegroundColor(blankSnapshot.style().fill()));
      assertEquals(ExcelBorderStyle.THIN, blankSnapshot.style().border().top().style());
      assertEquals(
          new ExcelHyperlink.Url("https://example.com/report"),
          blankSnapshot.metadata().hyperlink().orElseThrow());
      assertEquals(
          new ExcelComment("Review", "GridGrind", false),
          blankSnapshot.metadata().comment().orElseThrow().toPlainComment());
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
              ExcelCellFill.patternForeground(ExcelFillPattern.SOLID, ExcelColor.rgb("#DDEBF7")),
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
      assertEquals(rgb("#DDEBF7"), fillForegroundColor(dateSnapshot.style().fill()));
      assertTrue(dateSnapshot.style().font().bold());
      assertTrue(dateSnapshot.style().alignment().wrapText());
      assertEquals(
          ExcelVerticalAlignment.TOP, dateSnapshot.style().alignment().verticalAlignment());
      assertEquals(ExcelBorderStyle.DOUBLE, dateSnapshot.style().border().right().style());

      assertEquals("yyyy-mm-dd hh:mm:ss", dateTimeSnapshot.style().numberFormat());
      assertEquals(rgb("#DDEBF7"), fillForegroundColor(dateTimeSnapshot.style().fill()));
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
                              Boolean.TRUE,
                              null,
                              null,
                              null,
                              ExcelColor.rgb("#FF0000"),
                              null,
                              null))))));
      sheet.applyStyle(
          "A1",
          new ExcelCellStyle(
              null,
              null,
              new ExcelCellFont(
                  null,
                  Boolean.TRUE,
                  "Aptos",
                  new ExcelFontHeight(260),
                  ExcelColor.rgb("#112233"),
                  null,
                  null),
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
      assertEquals(rgb("#112233"), firstRun.font().fontColor());
      assertFalse(firstRun.font().bold());
      assertTrue(firstRun.font().italic());

      ExcelRichTextRunSnapshot secondRun = textSnapshot.richText().runs().get(1);
      assertEquals(" Report", secondRun.text());
      assertEquals("Aptos", secondRun.font().fontName());
      assertEquals(new ExcelFontHeight(260), secondRun.font().fontHeight());
      assertEquals(rgb("#FF0000"), secondRun.font().fontColor());
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
                  ExcelColor.rgb("#A3A3A3"),
                  null,
                  Boolean.TRUE),
              ExcelCellFill.patternForeground(ExcelFillPattern.SOLID, ExcelColor.rgb("#CDCDCD")),
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

      assertEquals(rgb("#CDCDCD"), fillForegroundColor(firstCell.style().fill()));
      assertEquals(rgb("#A3A3A3"), firstCell.style().font().fontColor());
      assertTrue(firstCell.style().font().strikeout());
      assertEquals(ExcelBorderStyle.DASH_DOT, firstCell.style().border().top().style());
      assertEquals(rgb("#CDCDCD"), fillForegroundColor(secondCell.style().fill()));
      assertEquals(ExcelBorderStyle.DASH_DOT, secondCell.style().border().top().style());
      assertEquals(rgb("#CDCDCD"), fillForegroundColor(untouchedStyledCell.style().fill()));
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
              ExcelCellFill.patternForeground(ExcelFillPattern.SOLID, ExcelColor.rgb("#603A79")),
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
      assertEquals(rgb("#603A79"), fillForegroundColor(appendedFirstCell.style().fill()));
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
}
