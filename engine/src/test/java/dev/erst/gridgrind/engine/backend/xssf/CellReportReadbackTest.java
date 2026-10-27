package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.table.SchemaColumnReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.CellReadFacet;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/**
 * Verifies canonical reports through actual workbook readback rather than retired snapshot
 * conversion.
 */
class CellReportReadbackTest {
  @Test
  void formulaAndValueFacetsStayIndependentAndMissingCellsRespectProjection() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      var cells = workbook.getOrCreateSheet("Data").cells();
      cells.setCell("A1", new ResolvedFormula("1+2"));
      var value =
          assertInstanceOf(
              CellReport.FormulaReport.class,
              cells.snapshotCell("A1", CellReadProjection.of(CellReadFacet.VALUE)));
      assertEquals(new CellValue.NumberValue(3), value.value().orElseThrow());
      assertTrue(value.formula().isEmpty());
      assertTrue(value.displayValue().isEmpty());
      var formula =
          assertInstanceOf(
              CellReport.FormulaReport.class,
              cells.snapshotCell("A1", CellReadProjection.of(CellReadFacet.FORMULA)));
      assertEquals("1+2", formula.formula().orElseThrow());
      assertTrue(formula.value().isEmpty());
      var format =
          assertInstanceOf(
              CellReport.FormulaReport.class,
              cells.snapshotCell("A1", CellReadProjection.of(CellReadFacet.FORMAT)));
      assertEquals("3", format.displayValue().orElseThrow());
      assertTrue(format.value().isEmpty());
      var blank =
          assertInstanceOf(
              CellReport.BlankReport.class,
              cells.snapshotCell("Z20", CellReadProjection.of(CellReadFacet.FORMULA)));
      assertTrue(blank.value().isEmpty());
      assertTrue(blank.style().isEmpty());
      assertTrue(blank.displayValue().isEmpty());
    }
  }

  @Test
  void metadataAndRichTextFacetsExposeTheirActualStoredFacts() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      var sheet = workbook.getOrCreateSheet("Data");
      sheet
          .cells()
          .setCell(
              "A1",
              new ResolvedRichText(
                  List.of(
                      new ResolvedRichTextRun("Hello", Optional.empty()),
                      new ResolvedRichTextRun(
                          " world",
                          Optional.of(
                              new CellFontPatch(
                                  Optional.of(true),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty()))))));
      sheet.links().setHyperlink("A1", new HyperlinkTarget.Url("https://example.com"));
      sheet.comments().setComment("A1", new ResolvedComment("Review", "Reader", true));
      var minimal = sheet.cells().snapshotCell("A1", CellReadProjection.defaults());
      assertTrue(minimal.hyperlink().isEmpty());
      assertTrue(minimal.comment().isEmpty());
      assertTrue(minimal.style().isEmpty());
      var full =
          assertInstanceOf(
              CellReport.TextReport.class,
              sheet
                  .cells()
                  .snapshotCell(
                      "A1",
                      CellReadProjection.of(
                          CellReadFacet.VALUE,
                          CellReadFacet.RICH_TEXT_RUNS,
                          CellReadFacet.HYPERLINK,
                          CellReadFacet.COMMENT)));
      assertEquals("Hello world", full.value().orElseThrow().value());
      assertEquals(
          List.of("Hello", " world"),
          full.runs().orElseThrow().stream()
              .map(dev.erst.gridgrind.core.model.style.RichTextRunReport::text)
              .toList());
      assertEquals(new HyperlinkTarget.Url("https://example.com"), full.hyperlink().orElseThrow());
      assertEquals("Review", full.comment().orElseThrow().text());
    }
  }

  @Test
  void schemaReadsTemporalKindsFromActualProjectedCellsAndFormulaValues() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      var sheet = workbook.getOrCreateSheet("Data");
      var poi = workbook.xssfWorkbook().getSheet("Data");
      var headers = poi.createRow(0);
      headers.createCell(0).setCellValue("Date");
      headers.createCell(1).setCellValue("Time");
      headers.createCell(2).setCellValue("DateTime");
      headers.createCell(3).setCellValue("Number");
      var values = poi.createRow(1);
      values.createCell(0).setCellValue(45839);
      values.createCell(1).setCellValue(0.5);
      values.createCell(2).setCellFormula("45839+0.5");
      values.createCell(3).setCellValue(42);
      var date = workbook.xssfWorkbook().createCellStyle();
      date.setDataFormat(workbook.xssfWorkbook().createDataFormat().getFormat("m/d/yyyy"));
      values.getCell(0).setCellStyle(date);
      var time = workbook.xssfWorkbook().createCellStyle();
      time.setDataFormat(workbook.xssfWorkbook().createDataFormat().getFormat("h:mm"));
      values.getCell(1).setCellStyle(time);
      var dateTime = workbook.xssfWorkbook().createCellStyle();
      dateTime.setDataFormat(workbook.xssfWorkbook().createDataFormat().getFormat("m/d/yyyy h:mm"));
      values.getCell(2).setCellStyle(dateTime);
      var schema =
          new WorkbookIntrospector()
              .sheetSchema(
                  workbook,
                  "Data",
                  "A1",
                  2,
                  4,
                  CellReadProjection.of(CellReadFacet.TEMPORAL),
                  false);
      assertEquals(
          List.of("DATE", "TIME", "DATE_TIME", "NUMBER"),
          schema.columns().stream().map(column -> column.dominantType().orElseThrow()).toList());
      assertEquals(
          List.of("Date", "Time", "DateTime", "Number"),
          schema.columns().stream().map(SchemaColumnReport::headerDisplayValue).toList());
      var numeric =
          new WorkbookIntrospector()
              .sheetSchema(workbook, "Data", "A1", 2, 4, CellReadProjection.defaults(), false);
      assertEquals(
          List.of("NUMBER", "NUMBER", "NUMBER", "NUMBER"),
          numeric.columns().stream().map(column -> column.dominantType().orElseThrow()).toList());
      assertEquals(
          42d,
          assertInstanceOf(CellReport.NumberReport.class, sheet.cells().snapshotCell("D2"))
              .value()
              .orElseThrow()
              .value());
    }
  }
}
