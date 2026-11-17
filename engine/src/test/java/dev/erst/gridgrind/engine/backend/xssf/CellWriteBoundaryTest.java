package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.engine.cells.CellErrors;
import org.apache.poi.ss.usermodel.FormulaError;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Verifies cell-write behavior at the POI boundary. */
class CellWriteBoundaryTest {
  @ParameterizedTest
  @ValueSource(strings = {"#CIRCULAR_REF!", "#FUNCTION_NOT_IMPLEMENTED!"})
  void rejectsEvaluationOnlyErrorsWithoutErasingAnExistingFormula(String literal) throws Exception {
    var observed = new CellValue.ErrorValue(literal);
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      var cells = workbook.getOrCreateSheet("Budget").cells();
      cells.setCell(
          new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"), new ResolvedFormula("1+1"));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              cells.setCell(
                  new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
                  new dev.erst.gridgrind.core.model.value.CellValueSpec.ErrorValue(
                      observed.value())));
      assertEquals(
          "1+1", workbook.xssfWorkbook().getSheet("Budget").getRow(0).getCell(0).getCellFormula());
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
          .setCell(
              new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
              new ResolvedRawFormula("LAMBDA(x,x+1)(A1)"));

      assertEquals(
          "LAMBDA(x,x+1)(A1)",
          workbook.xssfWorkbook().getSheet("Budget").getRow(0).getCell(0).getCellFormula());
    }
  }
}
