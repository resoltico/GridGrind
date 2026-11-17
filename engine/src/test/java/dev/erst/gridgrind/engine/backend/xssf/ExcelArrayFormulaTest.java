package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.value.CellReport;
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

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Calc");
      seedSourceData(sheet);

      sheet
          .cells()
          .setArrayFormula(
              new dev.erst.gridgrind.core.vocabulary.RangeRef("D2:D4"),
              new ResolvedArrayFormula("B2:B4*C2:C4"));
      sheet
          .cells()
          .setArrayFormula(
              new dev.erst.gridgrind.core.vocabulary.RangeRef("F2"),
              new ResolvedArrayFormula("SUM(B2:C2)"));

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
          assertInstanceOf(
              CellReport.FormulaReport.class,
              sheet.cells().snapshotCell(new dev.erst.gridgrind.core.vocabulary.CellAddress("D3")));
      assertEquals("B2:B4*C2:C4", formulaCell.formula().orElseThrow());

      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XssfWorkbook reopened =
        XssfWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      XssfSheet sheet = reopened.sheet("Calc");
      assertEquals(2, sheet.cells().arrayFormulas().size());
      assertEquals(
          List.of("D2:D4", "F2"),
          sheet.cells().arrayFormulas().stream()
              .map(report -> report.range().value())
              .sorted()
              .toList());

      sheet.cells().clearArrayFormula(new dev.erst.gridgrind.core.vocabulary.CellAddress("D3"));
      assertEquals(
          List.of("F2"),
          sheet.cells().arrayFormulas().stream().map(report -> report.range().value()).toList());
      assertEquals(
          "BLANK",
          sheet
              .cells()
              .snapshotCell(new dev.erst.gridgrind.core.vocabulary.CellAddress("D2"))
              .type());

      IllegalArgumentException notArray =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet
                      .cells()
                      .clearArrayFormula(new dev.erst.gridgrind.core.vocabulary.CellAddress("A1")));
      assertTrue(notArray.getMessage().contains("is not part of an array formula"));
    }
  }

  @Test
  void arrayFormulaGroupsAreVisibleThroughWorkbookIntrospection() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      WorkbookIntrospector introspector = new WorkbookIntrospector();
      XssfSheet sheet = workbook.getOrCreateSheet("Calc");
      seedSourceData(sheet);
      sheet
          .cells()
          .setArrayFormula(
              new dev.erst.gridgrind.core.vocabulary.RangeRef("D2:D4"),
              new ResolvedArrayFormula("B2:B4*C2:C4"));

      dev.erst.gridgrind.core.plan.SheetInspectionResult.ArrayFormulasResult result =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.ArrayFormulasResult.class,
              introspector.execute(
                  workbook,
                  new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                      new StepId("array-formulas"),
                      new dev.erst.gridgrind.core.plan.SheetSelector.ByNames(
                          List.of("Calc").stream()
                              .map(dev.erst.gridgrind.core.vocabulary.SheetName::new)
                              .toList()),
                      new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                          .GetArrayFormulas())));
      assertEquals("array-formulas", result.stepId().value());
      assertEquals(1, result.arrayFormulas().size());
      assertEquals("Calc", result.arrayFormulas().getFirst().sheetName().value());
      assertEquals("D2:D4", result.arrayFormulas().getFirst().range().value());
    }
  }

  private static void seedSourceData(XssfSheet sheet) {
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Month"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("B1"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Plan"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("C1"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Actual"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A2"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Jan"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A3"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Feb"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A4"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Mar"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("B2"),
            new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(10d));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("B3"),
            new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(18d));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("B4"),
            new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(15d));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("C2"),
            new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(12d));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("C3"),
            new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(16d));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("C4"),
            new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(21d));
  }
}
