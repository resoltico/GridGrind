package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.engine.cells.CellErrors;
import org.apache.poi.ss.usermodel.FormulaError;
import org.junit.jupiter.api.Test;

/** Verifies cell-write behavior at the POI boundary. */
class CellWriteBoundaryTest {
  @Test
  void rejectsEvaluationOnlyErrorsWithoutErasingAnExistingFormula() throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      var cells = workbook.getOrCreateSheet("Budget").cells();
      cells.setCell("A1", new ResolvedFormula("1+1"));
      for (CellValue.ErrorValue value :
          java.util.List.of(
              new CellValue.ErrorValue("#CIRCULAR_REF!"),
              new CellValue.ErrorValue("#FUNCTION_NOT_IMPLEMENTED!"))) {
        assertThrows(IllegalArgumentException.class, () -> cells.setCell("A1", value));
        assertEquals(
            "1+1",
            workbook.xssfWorkbook().getSheet("Budget").getRow(0).getCell(0).getCellFormula());
      }
    }
  }

  @Test
  void rejectsPoiFormulaErrorsThatHaveNoPublishedGridGrindWireLiteral() {
    IllegalArgumentException unsupported =
        assertThrows(
            IllegalArgumentException.class,
            () -> CellErrors.toReportedWireValue(FormulaError._NO_ERROR));

    assertEquals("_NO_ERROR is not a publishable cell error", unsupported.getMessage());
  }

  @Test
  void writesOpaqueFormulaValuesThroughTheSheetMutationBoundary() throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Budget")
          .cells()
          .setCell("A1", new ResolvedRawFormula("LAMBDA(x,x+1)(A1)"));

      assertEquals(
          "LAMBDA(x,x+1)(A1)",
          workbook.xssfWorkbook().getSheet("Budget").getRow(0).getCell(0).getCellFormula());
    }
  }
}
