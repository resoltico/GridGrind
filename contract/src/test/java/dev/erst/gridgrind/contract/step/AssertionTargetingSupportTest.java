package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for assertion-targeting helper seams extracted from workbook-step validation. */
class AssertionTargetingSupportTest {
  private static final Map<Class<? extends Assertion>, Class<? extends Selector>[]> STATIC_TYPES =
      Map.of(
          Assertion.TablePresent.class, targetTypes(TableSelector.ByNameOnSheet.class),
          Assertion.TableAbsent.class, targetTypes(TableSelector.ByNameOnSheet.class),
          Assertion.CellValue.class, targetTypes(CellSelector.ByAddress.class));

  @Test
  void resolvesAllowedTargetTypesAcrossDirectNestedAndCompositeAssertions() {
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            AssertionTargetingSupport.allowedTargetTypes(
                new Assertion.AnalysisMaxSeverity(
                    new InspectionQuery.AnalyzeFormulaHealth(), AnalysisSeverity.WARNING),
                STATIC_TYPES)));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            AssertionTargetingSupport.allowedTargetTypes(
                new Assertion.AnalysisFindingPresent(
                    new InspectionQuery.AnalyzeFormulaHealth(),
                    AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    null,
                    null),
                STATIC_TYPES)));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            AssertionTargetingSupport.allowedTargetTypes(
                new Assertion.AnalysisFindingAbsent(
                    new InspectionQuery.AnalyzeFormulaHealth(),
                    AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    null,
                    null),
                STATIC_TYPES)));
    assertEquals(
        List.of(TableSelector.ByNameOnSheet.class),
        List.of(
            AssertionTargetingSupport.allowedTargetTypes(
                new Assertion.AllOf(
                    List.of(new Assertion.TablePresent(), new Assertion.TableAbsent())),
                STATIC_TYPES)));
    assertEquals(
        List.of(TableSelector.ByNameOnSheet.class),
        List.of(
            AssertionTargetingSupport.allowedTargetTypes(
                new Assertion.AnyOf(
                    List.of(new Assertion.TablePresent(), new Assertion.TableAbsent())),
                STATIC_TYPES)));
    assertEquals(
        List.of(TableSelector.ByNameOnSheet.class),
        List.of(
            AssertionTargetingSupport.allowedTargetTypes(
                new Assertion.Not(new Assertion.TablePresent()), STATIC_TYPES)));
    assertEquals(
        List.of(CellSelector.ByAddress.class),
        List.of(
            AssertionTargetingSupport.allowedTargetTypes(
                new Assertion.CellValue(new ExpectedCellValue.Text("Owner")), STATIC_TYPES)));
  }

  @Test
  void exposesDynamicRuleTextAndStaticMappings() {
    assertEquals(
        Optional.of("Matches the nested analysis query's target selectors."),
        AssertionTargetingSupport.dynamicTargetSelectorRuleForAssertionType(
            Assertion.AnalysisMaxSeverity.class));
    assertEquals(
        Optional.of("Matches the intersection of every nested assertion's target selectors."),
        AssertionTargetingSupport.dynamicTargetSelectorRuleForAssertionType(Assertion.AllOf.class));
    assertEquals(
        Optional.of("Matches the nested assertion's target selectors."),
        AssertionTargetingSupport.targetSelectorRuleForAssertionType(Assertion.Not.class));
    assertEquals(
        Optional.empty(),
        AssertionTargetingSupport.targetSelectorRuleForAssertionType(Assertion.TablePresent.class));
    assertEquals(
        List.of(TableSelector.ByNameOnSheet.class),
        List.of(
            AssertionTargetingSupport.staticAllowedTargetTypesForAssertionType(
                Assertion.TablePresent.class, STATIC_TYPES)));
  }

  @Test
  void rejectsDynamicMissingAndIncompatibleCompositeMappings() {
    IllegalArgumentException dynamicFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                AssertionTargetingSupport.staticAllowedTargetTypesForAssertionType(
                    Assertion.Not.class, STATIC_TYPES));
    assertEquals(
        "Assertion type dev.erst.gridgrind.contract.assertion.Assertion$Not derives target selectors dynamically: Matches the nested assertion's target selectors.",
        dynamicFailure.getMessage());

    IllegalArgumentException missingFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                AssertionTargetingSupport.staticAllowedTargetTypesForAssertionType(
                    Assertion.ChartPresent.class, STATIC_TYPES));
    assertEquals(
        "No target-type mapping configured for assertion class dev.erst.gridgrind.contract.assertion.Assertion$ChartPresent",
        missingFailure.getMessage());

    IllegalArgumentException emptyFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> AssertionTargetingSupport.commonTargetTypes(List.of(), "ALL_OF", STATIC_TYPES));
    assertEquals(
        "ALL_OF requires nested assertions with compatible target families",
        emptyFailure.getMessage());

    IllegalArgumentException disjointFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                AssertionTargetingSupport.commonTargetTypes(
                    List.of(
                        new Assertion.TablePresent(),
                        new Assertion.CellValue(new ExpectedCellValue.Text("Owner"))),
                    "ALL_OF",
                    STATIC_TYPES));
    assertEquals(
        "ALL_OF requires nested assertions with compatible target families",
        disjointFailure.getMessage());
  }

  @SafeVarargs
  private static Class<? extends Selector>[] targetTypes(Class<? extends Selector>... types) {
    return types;
  }
}
