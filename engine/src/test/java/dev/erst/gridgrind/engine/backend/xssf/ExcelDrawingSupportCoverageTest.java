package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import dev.erst.gridgrind.engine.charts.ChartAdapter;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused residual-coverage tests for drawing support seam helpers. */
class ExcelDrawingSupportCoverageTest {
  @Test
  void supportHelperResidualBranchesStayCovered() throws Exception {
    assertTrue(DrawingAdapter.blankAsOptional(null).isEmpty());
    assertTrue(ChartAdapter.blankAsOptional(null).isEmpty());

    IllegalArgumentException blankFailure =
        assertThrows(
            IllegalArgumentException.class, () -> TextValidation.requireNonBlank(" ", "formula"));
    assertEquals("formula must not be blank", blankFailure.getMessage());

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Ops");
      var definedName = workbook.createName();
      definedName.setNameName("Source");
      definedName.setRefersToFormula("Ops!$A$1");
      assertEquals("Ops!$A$1", ChartAdapter.requiredDefinedNameFormula(definedName));
    }
  }
}
