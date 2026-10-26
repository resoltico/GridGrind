package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** End-to-end coverage for POI-backed array-formula authoring and readback. */
class ExcelArrayFormulaTest {
  @Test
  void arrayFormulaGroupsRoundTripThroughDirectSheetApis() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-array-formula-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Calc");
      seedSourceData(sheet);

      sheet.cells().setArrayFormula("D2:D4", new ResolvedArrayFormula("B2:B4*C2:C4"));
      sheet.cells().setArrayFormula("F2", new ResolvedArrayFormula("SUM(B2:C2)"));

      ArrayFormulaReport multiCell =
          sheet.cells().arrayFormulas().stream()
              .filter(snapshot -> "D2:D4".equals(snapshot.range().value()))
              .findFirst()
              .orElseThrow();
      ArrayFormulaReport singleCell =
          sheet.cells().arrayFormulas().stream()
              .filter(snapshot -> "F2".equals(snapshot.range().value()))
              .findFirst()
              .orElseThrow();
      assertEquals("D2", multiCell.topLeftAddress().value());
      assertEquals("B2:B4*C2:C4", multiCell.formula());
      assertFalse(multiCell.singleCell());
      assertTrue(singleCell.singleCell());

      CellReport.FormulaReport formulaCell =
          assertInstanceOf(CellReport.FormulaReport.class, sheet.cells().snapshotCell("D3"));
      assertEquals("B2:B4*C2:C4", formulaCell.formula().orElseThrow());

      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelSheet sheet = reopened.sheet("Calc");
      assertEquals(2, sheet.cells().arrayFormulas().size());
      assertEquals(
          List.of("D2:D4", "F2"),
          sheet.cells().arrayFormulas().stream()
              .map(report -> report.range().value())
              .sorted()
              .toList());

      sheet.cells().clearArrayFormula("D3");
      assertEquals(
          List.of("F2"),
          sheet.cells().arrayFormulas().stream().map(report -> report.range().value()).toList());
      assertEquals("BLANK", sheet.cells().snapshotCell("D2").type());

      IllegalArgumentException notArray =
          assertThrows(IllegalArgumentException.class, () -> sheet.cells().clearArrayFormula("A1"));
      assertTrue(notArray.getMessage().contains("is not part of an array formula"));
    }
  }

  @Test
  void arrayFormulaGroupsAreVisibleThroughWorkbookIntrospection() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();
      ExcelSheet sheet = workbook.getOrCreateSheet("Calc");
      seedSourceData(sheet);
      sheet.cells().setArrayFormula("D2:D4", new ResolvedArrayFormula("B2:B4*C2:C4"));

      dev.erst.gridgrind.core.plan.SheetInspectionResult.ArrayFormulasResult result =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.ArrayFormulasResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetArrayFormulas(
                      new StepId("array-formulas"),
                      new ExcelSheetSelection.Selected(List.of("Calc")))));
      assertEquals("array-formulas", result.stepId().value());
      assertEquals(1, result.arrayFormulas().size());
      assertEquals("Calc", result.arrayFormulas().getFirst().sheetName().value());
      assertEquals("D2:D4", result.arrayFormulas().getFirst().range().value());
    }
  }

  private static void seedSourceData(ExcelSheet sheet) {
    sheet.cells().setCell("A1", new CellValue.Text("Month"));
    sheet.cells().setCell("B1", new CellValue.Text("Plan"));
    sheet.cells().setCell("C1", new CellValue.Text("Actual"));
    sheet.cells().setCell("A2", new CellValue.Text("Jan"));
    sheet.cells().setCell("A3", new CellValue.Text("Feb"));
    sheet.cells().setCell("A4", new CellValue.Text("Mar"));
    sheet.cells().setCell("B2", new CellValue.NumberValue(10d));
    sheet.cells().setCell("B3", new CellValue.NumberValue(18d));
    sheet.cells().setCell("B4", new CellValue.NumberValue(15d));
    sheet.cells().setCell("C2", new CellValue.NumberValue(12d));
    sheet.cells().setCell("C3", new CellValue.NumberValue(16d));
    sheet.cells().setCell("C4", new CellValue.NumberValue(21d));
  }
}
