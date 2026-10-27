package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.model.formula.FormulaCapabilityKind;
import dev.erst.gridgrind.core.model.formula.FormulaCellTarget;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.engine.formulas.FormulaAdapter;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Direct coverage for the workbook formula facade. */
class ExcelWorkbookFormulasTest {
  @Test
  void wrapRejectsNullPoiWorkbook() {
    assertThrows(NullPointerException.class, () -> XssfWorkbooks.wrap(null));
  }

  @Test
  void wrapAdaptsMaterializedPoiWorkbook() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        XssfWorkbook workbook = XssfWorkbooks.wrap(poiWorkbook)) {
      poiWorkbook.createSheet("Budget");

      assertEquals(List.of("Budget"), workbook.sheets().sheetNames());
    }
  }

  @Test
  void formulasSurfaceDelegatesToWorkbookFormulaOperations() throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").cells().setCell("A1", new CellValue.NumberValue(2.0d));
      workbook.sheet("Budget").cells().setCell("B1", new ResolvedFormula("A1*2"));

      FormulaAdapter formulas = workbook.formulas();

      assertFalse(formulas.recalculateOnOpenEnabled());
      formulas.markRecalculateOnOpen();
      assertTrue(formulas.recalculateOnOpenEnabled());
      formulas.evaluateAll();
      assertEquals(
          4.0d,
          ((dev.erst.gridgrind.core.model.value.CellValue.NumberValue)
                  ((CellReport.FormulaReport) workbook.sheet("Budget").cells().snapshotCell("B1"))
                      .value()
                      .orElseThrow())
              .value());
      assertEquals(
          List.of(
              new CalculationReport.FormulaCapability(
                  new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")),
                  "A1*2",
                  FormulaCapabilityKind.EVALUABLE_NOW,
                  java.util.Optional.empty(),
                  java.util.Optional.empty())),
          formulas.assessAllCapabilities());
      assertEquals(
          List.of(
              new CalculationReport.FormulaCapability(
                  new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")),
                  "A1*2",
                  FormulaCapabilityKind.EVALUABLE_NOW,
                  java.util.Optional.empty(),
                  java.util.Optional.empty())),
          formulas.assessCapabilities(
              List.of(new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")))));
      formulas.clearCaches();
    }
  }

  @Test
  void classifiesSelfAndMutualCircularReferencesAsUnevaluableWithoutReclassifyingOrdinaryErrors()
      throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Budget");
      sheet.cells().setCell("A1", new ResolvedFormula("A1+1"));
      sheet.cells().setCell("B1", new ResolvedFormula("C1+1"));
      sheet.cells().setCell("C1", new ResolvedFormula("B1+1"));
      sheet.cells().setCell("D1", new ResolvedFormula("1/0"));

      List<CalculationReport.FormulaCapability> assessments =
          workbook.formulas().assessAllCapabilities();

      assertEquals(
          List.of(
              new CalculationReport.FormulaCapability(
                  new FormulaCellTarget(new SheetName("Budget"), new CellAddress("A1")),
                  "A1+1",
                  FormulaCapabilityKind.UNEVALUABLE_NOW,
                  java.util.Optional.of(GridGrindProblemCode.CIRCULAR_FORMULA_REFERENCE),
                  java.util.Optional.of("Formula evaluation detected a circular reference.")),
              new CalculationReport.FormulaCapability(
                  new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")),
                  "C1+1",
                  FormulaCapabilityKind.UNEVALUABLE_NOW,
                  java.util.Optional.of(GridGrindProblemCode.CIRCULAR_FORMULA_REFERENCE),
                  java.util.Optional.of("Formula evaluation detected a circular reference.")),
              new CalculationReport.FormulaCapability(
                  new FormulaCellTarget(new SheetName("Budget"), new CellAddress("C1")),
                  "B1+1",
                  FormulaCapabilityKind.UNEVALUABLE_NOW,
                  java.util.Optional.of(GridGrindProblemCode.CIRCULAR_FORMULA_REFERENCE),
                  java.util.Optional.of("Formula evaluation detected a circular reference.")),
              new CalculationReport.FormulaCapability(
                  new FormulaCellTarget(new SheetName("Budget"), new CellAddress("D1")),
                  "1/0",
                  FormulaCapabilityKind.EVALUABLE_NOW,
                  java.util.Optional.empty(),
                  java.util.Optional.empty())),
          assessments);
    }
  }

  @Test
  void treatsAnEvaluatorWithoutACellValueAsNonCircular() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        XssfWorkbook workbook =
            new XssfWorkbook(
                poiWorkbook,
                FormulaRuntimeTestDouble.nullEvaluation(
                    poiWorkbook.getCreationHelper().createFormulaEvaluator()))) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new ResolvedFormula("1+1"));

      assertEquals(
          List.of(
              new CalculationReport.FormulaCapability(
                  new FormulaCellTarget(new SheetName("Budget"), new CellAddress("A1")),
                  "1+1",
                  FormulaCapabilityKind.EVALUABLE_NOW,
                  java.util.Optional.empty(),
                  java.util.Optional.empty())),
          workbook.formulas().assessAllCapabilities());
    }
  }
}
