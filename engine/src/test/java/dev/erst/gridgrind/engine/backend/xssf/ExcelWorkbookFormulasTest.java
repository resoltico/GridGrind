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

  @Test
  void negationPrecedesExponentiationAsInExcel() throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      var cells = workbook.getOrCreateSheet("Arithmetic").cells();
      cells.setCell("A1", new ResolvedFormula("-2^2"));
      cells.setCell("A2", new ResolvedFormula("-(2^2)"));
      assertEquals(4.0d, cells.number("A1"));
      assertEquals(-4.0d, cells.number("A2"));
    }
  }

  @Test
  void multiplyAndDividePreserveBinaryPrecisionThroughWorkbookRoundtrip() throws Exception {
    // IEEE-754 binary64 results, independent of decimal display rounding or evaluator code.
    long multiplyBits = 0x4105aae2d70a3d70L;
    long divideBits = 0x3f9eb851eb851ebaL;
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      var cells = workbook.getOrCreateSheet("Precision").cells();
      cells.setCell("B2", new CellValue.NumberValue(125000.25));
      cells.setCell("B3", new CellValue.NumberValue(0.42));
      cells.setCell("B4", new ResolvedFormula("B2*(1+B3)"));
      cells.setCell("B5", new ResolvedFormula("(0.1+0.2)/10"));
      workbook.formulas().evaluateAll();
      assertEquals(multiplyBits, Double.doubleToRawLongBits(cells.number("B4")));
      assertEquals(divideBits, Double.doubleToRawLongBits(cells.number("B5")));
      var bytes = new java.io.ByteArrayOutputStream();
      workbook.xssfWorkbook().write(bytes);
      try (XssfWorkbook reopened =
          XssfWorkbooks.wrap(
              new XSSFWorkbook(new java.io.ByteArrayInputStream(bytes.toByteArray())))) {
        var persisted = reopened.xssfWorkbook().getSheetAt(0);
        assertEquals(
            multiplyBits,
            Double.doubleToRawLongBits(persisted.getRow(3).getCell(1).getNumericCellValue()));
        assertEquals(
            divideBits,
            Double.doubleToRawLongBits(persisted.getRow(4).getCell(1).getNumericCellValue()));
        var displayed =
            (CellReport.FormulaReport)
                reopened
                    .sheet("Precision")
                    .cells()
                    .snapshotCell(
                        "B4",
                        dev.erst.gridgrind.core.plan.CellReadProjection.of(
                            dev.erst.gridgrind.core.plan.CellReadFacet.FORMAT));
        assertEquals("177500.355", displayed.displayValue().orElseThrow());
      }
    }
  }
}
