package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Direct coverage for the workbook formula facade. */
class ExcelWorkbookFormulasTest {
  @Test
  void constructorRejectsNullWorkbook() {
    assertThrows(NullPointerException.class, () -> new ExcelWorkbookFormulas(null));
  }

  @Test
  void wrapRejectsNullPoiWorkbook() {
    assertThrows(NullPointerException.class, () -> ExcelWorkbooks.wrap(null));
  }

  @Test
  void wrapAdaptsMaterializedPoiWorkbook() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        ExcelWorkbook workbook = ExcelWorkbooks.wrap(poiWorkbook)) {
      poiWorkbook.createSheet("Budget");

      assertEquals(List.of("Budget"), workbook.sheets().sheetNames());
    }
  }

  @Test
  void formulasSurfaceDelegatesToWorkbookFormulaOperations() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").cells().setCell("A1", ExcelCellValue.number(2.0d));
      workbook.sheet("Budget").cells().setCell("B1", ExcelCellValue.formula("A1*2"));

      ExcelWorkbookFormulas formulas = workbook.formulas();

      assertFalse(formulas.recalculateOnOpenEnabled());
      assertSame(workbook, formulas.markRecalculateOnOpen());
      assertTrue(formulas.recalculateOnOpenEnabled());
      assertSame(workbook, formulas.evaluateAll());
      assertEquals(
          4.0d,
          ((ExcelCellSnapshot.NumberSnapshot)
                  ((ExcelCellSnapshot.FormulaSnapshot)
                          workbook.sheet("Budget").cells().snapshotCell("B1"))
                      .evaluation())
              .numberValue());
      assertEquals(
          List.of(
              new ExcelFormulaCapabilityAssessment(
                  "Budget", "B1", "A1*2", ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null)),
          formulas.assessAllCapabilities());
      assertEquals(
          List.of(
              new ExcelFormulaCapabilityAssessment(
                  "Budget", "B1", "A1*2", ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null)),
          formulas.assessCapabilities(List.of(new ExcelFormulaCellTarget("Budget", "B1"))));
      assertSame(workbook, formulas.clearCaches());
    }
  }
}
