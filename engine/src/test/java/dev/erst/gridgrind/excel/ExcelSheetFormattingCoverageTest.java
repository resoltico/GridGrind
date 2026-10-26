package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;
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

      sheet.cells().setCell("A1", new CellValue.Text("Header"));
      sheet.annotations().setHyperlink("A4", new HyperlinkTarget.Document("Budget!A1"));
      sheet.annotations().setComment("B5", new ResolvedComment("Note", "GridGrind", false));
      sheet.cells().appendRow(new CellValue.Text("Hosting"), new CellValue.NumberValue(49.0));

      assertEquals("Hosting", sheet.cells().text("A2"));
      assertEquals(49.0, sheet.cells().number("B2"));
      assertEquals(
          new HyperlinkTarget.Document("Budget!A1"),
          sheet.cells().snapshotCell("A4").hyperlink().orElseThrow());
      assertAll(
          () ->
              assertEquals("Note", sheet.cells().snapshotCell("B5").comment().orElseThrow().text()),
          () ->
              assertEquals(
                  "GridGrind", sheet.cells().snapshotCell("B5").comment().orElseThrow().author()),
          () -> assertFalse(sheet.cells().snapshotCell("B5").comment().orElseThrow().visible()));
    }
  }

  @Test
  void setCellPreservesExistingStyleAndMetadataOnStyledBlankCells() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet
          .cells()
          .applyStyle(
              "A1",
              new CellStylePatch(
                  Optional.empty(),
                  Optional.of(
                      new CellAlignmentPatch(
                          Optional.empty(),
                          Optional.of(ExcelHorizontalAlignment.CENTER),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.of(
                      new CellFontPatch(
                          Optional.of(Boolean.TRUE),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.of(
                      CellFillSpec.patternForeground(
                          ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#AABBCC")))),
                  Optional.of(
                      new CellBorderPatch(
                          Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.THIN)),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty()));
      sheet.annotations().setHyperlink("A1", new HyperlinkTarget.Url("https://example.com/report"));
      sheet.annotations().setComment("A1", new ResolvedComment("Review", "GridGrind", false));

      CellReport.TextReport textSnapshot =
          (CellReport.TextReport)
              sheet
                  .cells()
                  .setCell("A1", new CellValue.Text("Quarterly report"))
                  .snapshotCell("A1");
      assertEquals("Quarterly report", textSnapshot.value().orElseThrow().value());
      assertEquals(rgb("#AABBCC"), fillForegroundColor(textSnapshot.style().orElseThrow().fill()));
      assertEquals(
          ExcelBorderStyle.THIN, borderStyle(textSnapshot.style().orElseThrow().border().top()));
      assertEquals(
          ExcelHorizontalAlignment.CENTER,
          textSnapshot.style().orElseThrow().alignment().horizontalAlignment());
      assertEquals(
          new HyperlinkTarget.Url("https://example.com/report"),
          textSnapshot.hyperlink().orElseThrow());
      assertAll(
          () -> assertEquals("Review", textSnapshot.comment().orElseThrow().text()),
          () -> assertEquals("GridGrind", textSnapshot.comment().orElseThrow().author()),
          () -> assertFalse(textSnapshot.comment().orElseThrow().visible()));

      CellReport.BlankReport blankSnapshot =
          (CellReport.BlankReport)
              sheet.cells().setCell("A1", new CellValue.Blank()).snapshotCell("A1");
      assertEquals(rgb("#AABBCC"), fillForegroundColor(blankSnapshot.style().orElseThrow().fill()));
      assertEquals(
          ExcelBorderStyle.THIN, borderStyle(blankSnapshot.style().orElseThrow().border().top()));
      assertEquals(
          new HyperlinkTarget.Url("https://example.com/report"),
          blankSnapshot.hyperlink().orElseThrow());
      assertAll(
          () -> assertEquals("Review", blankSnapshot.comment().orElseThrow().text()),
          () -> assertEquals("GridGrind", blankSnapshot.comment().orElseThrow().author()),
          () -> assertFalse(blankSnapshot.comment().orElseThrow().visible()));
    }
  }

  @Test
  void dateValueWritesMergeRequiredNumberFormatsOntoExistingStyle() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      CellStylePatch style =
          new CellStylePatch(
              Optional.empty(),
              Optional.of(
                  new CellAlignmentPatch(
                      Optional.of(Boolean.TRUE),
                      Optional.empty(),
                      Optional.of(ExcelVerticalAlignment.TOP),
                      Optional.empty(),
                      Optional.empty())),
              Optional.of(
                  new CellFontPatch(
                      Optional.of(Boolean.TRUE),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty())),
              Optional.of(
                  CellFillSpec.patternForeground(
                      ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#DDEBF7")))),
              Optional.of(
                  new CellBorderPatch(
                      Optional.empty(),
                      Optional.empty(),
                      Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.DOUBLE)),
                      Optional.empty(),
                      Optional.empty())),
              Optional.empty());
      sheet.cells().applyStyle("A1:B1", style);

      CellReport.NumberReport dateSnapshot =
          (CellReport.NumberReport)
              sheet
                  .cells()
                  .setCell("A1", new CellValueSpec.Date(LocalDate.of(2026, 3, 31)))
                  .snapshotCell("A1");
      CellReport.NumberReport dateTimeSnapshot =
          (CellReport.NumberReport)
              sheet
                  .cells()
                  .setCell(
                      "B1", new CellValueSpec.DateTime(LocalDateTime.of(2026, 3, 31, 9, 45, 0)))
                  .snapshotCell("B1");

      assertEquals("yyyy-mm-dd", dateSnapshot.style().orElseThrow().numberFormat());
      assertEquals(rgb("#DDEBF7"), fillForegroundColor(dateSnapshot.style().orElseThrow().fill()));
      assertTrue(dateSnapshot.style().orElseThrow().font().bold());
      assertTrue(dateSnapshot.style().orElseThrow().alignment().wrapText());
      assertEquals(
          ExcelVerticalAlignment.TOP,
          dateSnapshot.style().orElseThrow().alignment().verticalAlignment());
      assertEquals(
          ExcelBorderStyle.DOUBLE,
          borderStyle(dateSnapshot.style().orElseThrow().border().right()));

      assertEquals("yyyy-mm-dd hh:mm:ss", dateTimeSnapshot.style().orElseThrow().numberFormat());
      assertEquals(
          rgb("#DDEBF7"), fillForegroundColor(dateTimeSnapshot.style().orElseThrow().fill()));
      assertTrue(dateTimeSnapshot.style().orElseThrow().font().bold());
      assertTrue(dateTimeSnapshot.style().orElseThrow().alignment().wrapText());
      assertEquals(
          ExcelVerticalAlignment.TOP,
          dateTimeSnapshot.style().orElseThrow().alignment().verticalAlignment());
      assertEquals(
          ExcelBorderStyle.DOUBLE,
          borderStyle(dateTimeSnapshot.style().orElseThrow().border().right()));
    }
  }

  @Test
  void snapshotsRichTextRunsWithEffectiveInheritedFontFacts() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet
          .cells()
          .setCell(
              "A1",
              new ResolvedRichText(
                  List.of(
                      new ResolvedRichTextRun("Quarterly", Optional.empty()),
                      new ResolvedRichTextRun(
                          " Report",
                          Optional.of(
                              new CellFontPatch(
                                  Optional.of(Boolean.TRUE),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.of(ColorSpec.rgb(new RgbHex("#FF0000"))),
                                  Optional.empty(),
                                  Optional.empty()))))));
      sheet
          .cells()
          .applyStyle(
              "A1",
              new CellStylePatch(
                  Optional.empty(),
                  Optional.empty(),
                  Optional.of(
                      new CellFontPatch(
                          Optional.empty(),
                          Optional.of(Boolean.TRUE),
                          Optional.of("Aptos"),
                          Optional.of(new FontHeightSpec.Twips(260)),
                          Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty(),
                  Optional.empty(),
                  Optional.empty()));

      CellReport.TextReport textSnapshot = (CellReport.TextReport) sheet.cells().snapshotCell("A1");

      assertEquals("Quarterly Report", textSnapshot.value().orElseThrow().value());
      assertNotNull(textSnapshot.runs());
      assertEquals(
          "Quarterly Report",
          textSnapshot.runs().orElseThrow().stream()
              .map(RichTextRunReport::text)
              .collect(java.util.stream.Collectors.joining()));
      assertEquals(2, textSnapshot.runs().orElseThrow().size());

      RichTextRunReport firstRun = textSnapshot.runs().orElseThrow().get(0);
      assertEquals("Quarterly", firstRun.text());
      assertEquals("Aptos", firstRun.font().fontName());
      assertEquals(FontHeightReport.fromTwips(260), firstRun.font().fontHeight());
      assertEquals(rgb("#112233"), firstRun.font().fontColor());
      assertFalse(firstRun.font().bold());
      assertTrue(firstRun.font().italic());

      RichTextRunReport secondRun = textSnapshot.runs().orElseThrow().get(1);
      assertEquals(" Report", secondRun.text());
      assertEquals("Aptos", secondRun.font().fontName());
      assertEquals(FontHeightReport.fromTwips(260), secondRun.font().fontHeight());
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

      sheet
          .cells()
          .applyStyle(
              "A1:B2",
              new CellStylePatch(
                  Optional.empty(),
                  Optional.empty(),
                  Optional.of(
                      new CellFontPatch(
                          Optional.empty(),
                          Optional.empty(),
                          Optional.of("Aptos"),
                          Optional.of(FontHeightReport.fromPoints(new java.math.BigDecimal("15.2")))
                              .map(height -> new FontHeightSpec.Twips(height.twips())),
                          Optional.of(ColorSpec.rgb(new RgbHex("#A3A3A3"))),
                          Optional.empty(),
                          Optional.of(Boolean.TRUE))),
                  Optional.of(
                      CellFillSpec.patternForeground(
                          ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#CDCDCD")))),
                  Optional.of(
                      new CellBorderPatch(
                          Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.DASH_DOT)),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty()));

      sheet
          .cells()
          .appendRow(
              new CellValue.NumberValue(607.8483822864587), new CellValue.BooleanValue(false));

      CellReport.NumberReport firstCell =
          (CellReport.NumberReport) sheet.cells().snapshotCell("A1");
      CellReport.BooleanReport secondCell =
          (CellReport.BooleanReport) sheet.cells().snapshotCell("B1");
      CellReport.BlankReport untouchedStyledCell =
          (CellReport.BlankReport) sheet.cells().snapshotCell("A2");

      assertEquals(rgb("#CDCDCD"), fillForegroundColor(firstCell.style().orElseThrow().fill()));
      assertEquals(rgb("#A3A3A3"), firstCell.style().orElseThrow().font().fontColor());
      assertTrue(firstCell.style().orElseThrow().font().strikeout());
      assertEquals(
          ExcelBorderStyle.DASH_DOT, borderStyle(firstCell.style().orElseThrow().border().top()));
      assertEquals(rgb("#CDCDCD"), fillForegroundColor(secondCell.style().orElseThrow().fill()));
      assertEquals(
          ExcelBorderStyle.DASH_DOT, borderStyle(secondCell.style().orElseThrow().border().top()));
      assertEquals(
          rgb("#CDCDCD"), fillForegroundColor(untouchedStyledCell.style().orElseThrow().fill()));
      assertEquals(
          ExcelBorderStyle.DASH_DOT,
          borderStyle(untouchedStyledCell.style().orElseThrow().border().top()));
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

      sheet
          .cells()
          .appendRow(
              new CellValueSpec.DateTime(LocalDateTime.of(2026, 2, 6, 13, 1, 1)),
              new CellValueSpec.DateTime(LocalDateTime.of(2026, 2, 6, 13, 58, 1)));
      sheet
          .cells()
          .applyStyle(
              "A1:B2",
              new CellStylePatch(
                  Optional.of("0.00"),
                  Optional.of(
                      new CellAlignmentPatch(
                          Optional.of(Boolean.TRUE),
                          Optional.empty(),
                          Optional.of(ExcelVerticalAlignment.CENTER),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.of(
                      new CellFontPatch(
                          Optional.of(Boolean.TRUE),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.of(
                      CellFillSpec.patternForeground(
                          ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#603A79")))),
                  Optional.of(
                      new CellBorderPatch(
                          Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.MEDIUM_DASHED)),
                          Optional.empty(),
                          Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.MEDIUM_DASHED)),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty()));

      sheet
          .cells()
          .appendRow(
              new CellValueSpec.DateTime(LocalDateTime.of(2026, 2, 6, 13, 1, 58)),
              new CellValueSpec.DateTime(LocalDateTime.of(2026, 2, 6, 13, 1, 1)));

      CellReport.NumberReport appendedFirstCell =
          (CellReport.NumberReport) sheet.cells().snapshotCell("A2");
      CellReport.NumberReport appendedSecondCell =
          (CellReport.NumberReport) sheet.cells().snapshotCell("B2");

      assertEquals("yyyy-mm-dd hh:mm:ss", appendedFirstCell.style().orElseThrow().numberFormat());
      assertEquals("yyyy-mm-dd hh:mm:ss", appendedSecondCell.style().orElseThrow().numberFormat());
      assertEquals(
          rgb("#603A79"), fillForegroundColor(appendedFirstCell.style().orElseThrow().fill()));
      assertTrue(appendedFirstCell.style().orElseThrow().font().bold());
      assertTrue(appendedFirstCell.style().orElseThrow().alignment().wrapText());
      assertEquals(
          ExcelVerticalAlignment.CENTER,
          appendedFirstCell.style().orElseThrow().alignment().verticalAlignment());
      assertEquals(
          ExcelBorderStyle.MEDIUM_DASHED,
          borderStyle(appendedFirstCell.style().orElseThrow().border().top()));
      assertEquals(
          ExcelBorderStyle.MEDIUM_DASHED,
          borderStyle(appendedFirstCell.style().orElseThrow().border().right()));
    }
  }

  @Test
  void appendRowWithNoValuesIsANoOp() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);
      ExcelSheetCells cells = sheet.cells();

      assertSame(cells, cells.appendRow());
      assertEquals(0, sheet.rows().physicalCount());
      assertEquals(-1, sheet.rows().lastIndex());
      assertEquals(-1, sheet.columns().lastIndex());
    }
  }
}
