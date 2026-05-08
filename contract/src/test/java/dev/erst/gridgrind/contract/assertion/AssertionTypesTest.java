package dev.erst.gridgrind.contract.assertion;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for the first-class assertion type system. */
class AssertionTypesTest {
  @Test
  void exposesStableDiscriminatorsAndValidatesLeafInputs() {
    assertEquals("EXPECT_TABLE_PRESENT", new PresenceAssertion.TablePresent().assertionType());
    assertEquals(
        "EXPECT_ANALYSIS_FINDING_PRESENT",
        new AnalysisAssertion.AnalysisFindingPresent(
                new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                AnalysisFindingCode.FORMULA_ERROR_RESULT,
                Optional.of(AnalysisSeverity.ERROR),
                Optional.of("error"))
            .assertionType());
    assertEquals(
        "ALL_OF",
        new CompositeAssertion.AllOf(List.of(new PresenceAssertion.TablePresent()))
            .assertionType());
    assertEquals(
        "expectedValue must not be null",
        assertThrows(NullPointerException.class, () -> new CellAssertion.CellValue(null))
            .getMessage());
    assertEquals(
        "messageContains must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new AnalysisAssertion.AnalysisFindingAbsent(
                        new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                        AnalysisFindingCode.FORMULA_ERROR_RESULT,
                        Optional.empty(),
                        Optional.of(" ")))
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
