package dev.erst.gridgrind.contract.assertion;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for the first-class assertion type system. */
class AssertionTypesTest {
  @Test
  void exposesStableDiscriminatorsAndValidatesLeafInputs() {
    assertEquals("EXPECT_TABLE_PRESENT", new Assertion.TablePresent().assertionType());
    assertEquals(
        "EXPECT_ANALYSIS_FINDING_PRESENT",
        new Assertion.AnalysisFindingPresent(
                new InspectionQuery.AnalyzeFormulaHealth(),
                AnalysisFindingCode.FORMULA_ERROR_RESULT,
                AnalysisSeverity.ERROR,
                "error")
            .assertionType());
    assertEquals(
        "ALL_OF", new Assertion.AllOf(List.of(new Assertion.TablePresent())).assertionType());
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
        new ExpectedCellValue.BooleanValue(true), new ExpectedCellValue.BooleanValue(true));
    assertEquals(
        "error must not be blank",
        assertThrows(IllegalArgumentException.class, () -> new ExpectedCellValue.ErrorValue(" "))
            .getMessage());
  }
}
