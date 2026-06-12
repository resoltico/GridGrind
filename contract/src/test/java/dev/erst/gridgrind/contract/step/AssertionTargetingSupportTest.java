package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
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
                    new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner")))));
  }

  @Test
  void exposesDynamicRuleTextAndStaticMappings() {
    assertEquals(
        Optional.of("Matches the nested analysis query's target selectors."),
        Assertion.dynamicTargetSelectorRuleForType(AnalysisAssertion.AnalysisMaxSeverity.class));
    assertEquals(
        Optional.of("Matches the intersection of every nested assertion's target selectors."),
        Assertion.dynamicTargetSelectorRuleForType(CompositeAssertion.AllOf.class));
    assertEquals(
        Optional.of("Matches the nested assertion's target selectors."),
        Assertion.targetSelectorRuleForType(CompositeAssertion.Not.class));
    assertEquals(
        Optional.empty(),
        Assertion.targetSelectorRuleForType(PresenceAssertion.TablePresent.class));
    assertEquals(
        List.of(TableSelector.class),
        List.of(Assertion.staticAllowedTargetTypesForType(PresenceAssertion.TablePresent.class)));
  }

  @Test
  void rejectsDynamicMissingAndIncompatibleCompositeMappings() {
    IllegalArgumentException dynamicFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> Assertion.staticAllowedTargetTypesForType(CompositeAssertion.Not.class));
    assertEquals(
        "Assertion type dev.erst.gridgrind.contract.assertion.CompositeAssertion$Not derives target selectors dynamically: Matches the nested assertion's target selectors.",
        dynamicFailure.getMessage());

    IllegalArgumentException disjointFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                Assertion.allowedTargetTypes(
                    new CompositeAssertion.AllOf(
                        List.of(
                            new PresenceAssertion.TablePresent(),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                    "Owner"))))));
    assertEquals(
        "ALL_OF requires nested assertions with compatible target families",
        disjointFailure.getMessage());
  }
}
