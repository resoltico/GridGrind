package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
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
        assertThrows(NullPointerException.class, () -> new CellAssertion.Value(null)).getMessage());
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
        "value must not be null",
        assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.Text(null))
            .getMessage());
    assertEquals(
        "value must be finite",
        assertThrows(
                IllegalArgumentException.class, () -> new CellValueSpec.NumberValue(Double.NaN))
            .getMessage());
    assertEquals(new CellValueSpec.BooleanValue(true), new CellValueSpec.BooleanValue(true));
    assertEquals(
        "value must not be blank",
        assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.ErrorValue(" "))
            .getMessage());
    assertEquals(
        "expectedValue must not be null",
        assertThrows(NullPointerException.class, () -> new CellAssertion.Value(null)).getMessage());
    assertEquals(new CellValue.BooleanValue(true), new CellValue.BooleanValue(true));
    assertEquals(
        "expectedValue.value must be finite",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellAssertion.Value(new CellValue.NumberValue(Double.POSITIVE_INFINITY)))
            .getMessage());
    assertEquals(
        "value must not be blank",
        assertThrows(IllegalArgumentException.class, () -> new CellValue.ErrorValue(" "))
            .getMessage());
  }
}
