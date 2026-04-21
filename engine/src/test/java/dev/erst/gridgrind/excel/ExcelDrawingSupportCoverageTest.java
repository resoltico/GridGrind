package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused residual-coverage tests for drawing support seam helpers. */
class ExcelDrawingSupportCoverageTest {
  @Test
  void supportHelperResidualBranchesStayCovered() throws Exception {
    assertNull(ExcelDrawingAnchorSupport.nullIfBlank(null));
    assertNull(ExcelChartSourceSupport.nullIfBlank(null));

    IllegalArgumentException blankFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelChartSourceSupport.requireNonBlank(" ", "formula"));
    assertEquals("formula must not be blank", blankFailure.getMessage());

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Ops");
      var definedName = workbook.createName();
      definedName.setNameName("Source");
      definedName.setRefersToFormula("Ops!$A$1");
      assertEquals("Ops!$A$1", ExcelDrawingChartSupport.requiredDefinedNameFormula(definedName));
    }
  }
}
