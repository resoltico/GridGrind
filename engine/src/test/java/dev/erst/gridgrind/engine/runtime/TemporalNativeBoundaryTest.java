package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.value.*;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.vocabulary.*;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.engine.cells.TemporalFormats;
import java.time.*;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Native guards use the shared owner before mutation; read facets retain independent facts. */
class TemporalNativeBoundaryTest {
  @Test
  void batchDateRefusalLeavesTheReceiving1904SheetUnchanged() throws Exception {
    try (var workbook = XssfWorkbooks.create()) {
      workbook.xssfWorkbook().getCTWorkbook().getWorkbookPr().setDate1904(true);
      var cells = workbook.getOrCreateSheet("Dates").cells();
      var nativeSheet = workbook.xssfWorkbook().getSheetAt(0);
      String before = nativeSheet.getCTWorksheet().xmlText();
      assertThrows(
          IllegalArgumentException.class,
          () ->
              cells.setRange(
                  new RangeRef("A1:B1"),
                  List.of(
                      List.of(
                          new CellValueSpec.Date(LocalDate.of(2026, 3, 8)),
                          new CellValueSpec.Date(LocalDate.of(1903, 12, 31))))));
      assertEquals(before, nativeSheet.getCTWorksheet().xmlText());
      assertThrows(
          IllegalArgumentException.class,
          () ->
              cells.appendRow(
                  new ResolvedCellText("otherwise mutable"),
                  new CellValueSpec.Date(LocalDate.of(1903, 12, 31))));
      assertEquals(before, nativeSheet.getCTWorksheet().xmlText());
      cells.setCell(new CellAddress("A1"), new CellValueSpec.Date(LocalDate.of(1904, 1, 1)));
      assertEquals(0, nativeSheet.getRow(0).getCell(0).getNumericCellValue());
      cells.setCell(
          new CellAddress("B1"),
          new CellValueSpec.DateTime(LocalDateTime.parse("2026-03-08T02:30:00")));
      assertEquals(44627.104166666664, nativeSheet.getRow(0).getCell(1).getNumericCellValue());
    }
  }

  @Test
  void optionalInterpretationDoesNotLoseNumbersOrUnrelatedCells() throws Exception {
    try (var workbook = XssfWorkbooks.create()) {
      var sheet = workbook.getOrCreateSheet("Observed");
      var nativeSheet = workbook.xssfWorkbook().getSheetAt(0);
      var row = nativeSheet.createRow(0);
      var style = workbook.xssfWorkbook().createCellStyle();
      style.setDataFormat(workbook.xssfWorkbook().createDataFormat().getFormat("yyyy-mm-dd"));
      row.createCell(0).setCellValue(-1);
      row.getCell(0).setCellStyle(style);
      row.createCell(1).setCellValue(60);
      row.getCell(1).setCellStyle(style);
      row.createCell(2).setCellValue(61);
      row.getCell(2).setCellStyle(style);
      var addresses = List.of(new CellAddress("A1"), new CellAddress("B1"), new CellAddress("C1"));
      var expectedNumbers =
          List.of(
              new CellValue.NumberValue(-1),
              new CellValue.NumberValue(60),
              new CellValue.NumberValue(61));
      var interpreted = new CellTemporalReport.Interpreted(CellTemporalKind.DATE, "1900-03-01");
      var projection =
          CellReadProjection.of(CellReadFacet.VALUE, CellReadFacet.FORMAT, CellReadFacet.TEMPORAL);
      for (int column = 0; column < 3; column++) {
        var report =
            assertInstanceOf(
                CellReport.NumberReport.class,
                sheet.cells().snapshotCell(addresses.get(column), projection));
        assertEquals(expectedNumbers.get(column), report.value().orElseThrow());
        assertTrue(report.displayValue().isPresent());
        if (column < 2)
          assertInstanceOf(
              CellTemporalReport.Unrepresentable.class, report.temporal().orElseThrow());
        else assertEquals(interpreted, report.temporal().orElseThrow());
      }
      var formula = row.createCell(3);
      formula.setCellFormula("C1");
      formula.setCellStyle(style);
      var formulaReport =
          assertInstanceOf(
              CellReport.FormulaReport.class,
              sheet
                  .cells()
                  .snapshotCell(
                      new CellAddress("D1"), CellReadProjection.of(CellReadFacet.TEMPORAL)));
      assertTrue(formulaReport.value().isEmpty());
      assertTrue(formulaReport.temporal().isPresent(), "TEMPORAL does not require the VALUE facet");
    }
  }

  @Test
  void elapsedClassificationDoesNotMistakeLiteralsColorsOrLocaleBrackets() {
    for (String value : List.of("[h]:mm", "[mm]:ss", "[ss]"))
      assertEquals(CellTemporalKind.ELAPSED, TemporalFormats.observedKind(value).orElseThrow());
    for (String value : List.of("\"[h]\" 0", "\\[h\\] 0", "[Red]0", "[$-409]0"))
      assertNotEquals(
          java.util.Optional.of(CellTemporalKind.ELAPSED), TemporalFormats.observedKind(value));
  }
}
