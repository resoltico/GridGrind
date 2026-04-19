package dev.erst.gridgrind.contract.assertion;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for the first-class assertion type system. */
class AssertionTypesTest {
  @Test
  void exposesStableDiscriminatorsAndValidatesLeafInputs() {
    assertEquals("EXPECT_PRESENT", new Assertion.Present().assertionType());
    assertEquals(
        "EXPECT_ANALYSIS_FINDING_PRESENT",
        new Assertion.AnalysisFindingPresent(
                new InspectionQuery.AnalyzeFormulaHealth(),
                AnalysisFindingCode.FORMULA_ERROR_RESULT,
                AnalysisSeverity.ERROR,
                "error")
            .assertionType());
    assertEquals("ALL_OF", new Assertion.AllOf(List.of(new Assertion.Present())).assertionType());
    assertEquals(
        "expectedValue must not be null",
        assertThrows(NullPointerException.class, () -> new Assertion.CellValue(null)).getMessage());
    assertEquals(
        "messageContains must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new Assertion.AnalysisFindingAbsent(
                        new InspectionQuery.AnalyzeFormulaHealth(),
                        AnalysisFindingCode.FORMULA_ERROR_RESULT,
                        null,
                        " "))
            .getMessage());
  }

  @Test
  void expectedCellValueVariantsValidateAndExposeWireShape() {
    assertEquals(
        "text must not be null",
        assertThrows(NullPointerException.class, () -> new ExpectedCellValue.Text(null))
            .getMessage());
    assertEquals(
        "number must be finite",
        assertThrows(
                IllegalArgumentException.class,
                () -> new ExpectedCellValue.NumericValue(Double.NaN))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(NullPointerException.class, () -> new ExpectedCellValue.BooleanValue(null))
            .getMessage());
    assertEquals(
        "error must not be blank",
        assertThrows(IllegalArgumentException.class, () -> new ExpectedCellValue.ErrorValue(" "))
            .getMessage());
  }
}
