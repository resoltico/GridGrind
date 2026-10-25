package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for assertion-owned selector targeting. */
class AssertionTargetingSupportTest {
  @Test
  void resolvesAllowedTargetTypesAcrossDirectNestedAndCompositeAssertions() {
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            Assertion.allowedTargetTypes(
                new AnalysisAssertion.AnalysisMaxSeverity(
                    new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                    AnalysisSeverity.WARNING))));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            Assertion.allowedTargetTypes(
                new AnalysisAssertion.AnalysisFindingPresent(
                    new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                    AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    Optional.empty(),
                    Optional.empty()))));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            Assertion.allowedTargetTypes(
                new AnalysisAssertion.AnalysisFindingAbsent(
                    new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                    AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    Optional.empty(),
                    Optional.empty()))));
    assertEquals(
        List.of(TableSelector.class),
        List.of(
            Assertion.allowedTargetTypes(
                new CompositeAssertion.AllOf(
                    List.of(
                        new PresenceAssertion.TablePresent(),
                        new PresenceAssertion.TableAbsent())))));
    assertEquals(
        List.of(TableSelector.class),
        List.of(
            Assertion.allowedTargetTypes(
                new CompositeAssertion.AnyOf(
                    List.of(
                        new PresenceAssertion.TablePresent(),
                        new PresenceAssertion.TableAbsent())))));
    assertEquals(
        List.of(TableSelector.class),
        List.of(
            Assertion.allowedTargetTypes(
                new CompositeAssertion.Not(new PresenceAssertion.TablePresent()))));
    assertEquals(
        List.of(
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(
            Assertion.allowedTargetTypes(
                new CellAssertion.CellValue(
                    new dev.erst.gridgrind.core.model.value.CellScalarValue.Text("Owner")))));
  }

  @Test
  void rejectsIncompatibleCompositeMappings() {
    IllegalArgumentException disjointFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                Assertion.allowedTargetTypes(
                    new CompositeAssertion.AllOf(
                        List.of(
                            new PresenceAssertion.TablePresent(),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.core.model.value.CellScalarValue.Text(
                                    "Owner"))))));
    assertEquals(
        "ALL_OF requires nested assertions with compatible target families",
        disjointFailure.getMessage());
  }
}
