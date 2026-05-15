package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import org.junit.jupiter.api.Test;

/** Tests for ordered inspection step validation and target compatibility. */
class InspectionStepTest {
  @Test
  void acceptsCompatibleTargetAndQueryPairs() {
    InspectionStep workbookSummary =
        new InspectionStep(
            "summary",
            new WorkbookSelector.Current(),
            new WorkbookIntrospectionQuery.GetWorkbookSummary());
    InspectionStep formulaHealth =
        new InspectionStep(
            "formula-health",
            new SheetSelector.ByNames(java.util.List.of("Budget")),
            new InspectionAnalysisQuery.AnalyzeFormulaHealth());

    assertEquals("INSPECTION", workbookSummary.stepKind());
    assertEquals("GET_WORKBOOK_SUMMARY", workbookSummary.query().queryType());
    assertEquals("ANALYZE_FORMULA_HEALTH", formulaHealth.query().queryType());
  }

  @Test
  void rejectsIncompatibleTargetsOrBlankStepIds() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new InspectionStep(
                " ",
                new WorkbookSelector.Current(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary()));
    IllegalArgumentException incompatibleTargetFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new InspectionStep(
                    "bad-target",
                    new RangeSelector.ByRange("Budget", "A1:B2"),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary()));

    assertEquals(
        "GET_WORKBOOK_SUMMARY requires target type WORKBOOK_CURRENT but got RANGE_BY_RANGE",
        incompatibleTargetFailure.getMessage());
  }
}
