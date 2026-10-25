package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.vocabulary.StepId;
import org.junit.jupiter.api.Test;

/** Tests for ordered inspection step validation and target compatibility. */
class InspectionStepTest {
  @Test
  void acceptsCompatibleTargetAndQueryPairs() {
    InspectionStep workbookSummary =
        new InspectionStep(
            new StepId("summary"),
            new WorkbookSelector.Current(),
            new WorkbookIntrospectionQuery.GetWorkbookSummary());
    InspectionStep formulaHealth =
        new InspectionStep(
            new StepId("formula-health"),
            new SheetSelector.ByNames(java.util.List.of("Budget")),
            new InspectionAnalysisQuery.AnalyzeFormulaHealth());

    assertEquals("INSPECTION", workbookSummary.stepKind());
    assertEquals("GET_WORKBOOK_SUMMARY", workbookSummary.query().queryType());
    assertEquals("ANALYZE_FORMULA_HEALTH", formulaHealth.query().queryType());
  }

  @Test
  void keepsIncompatibleTargetsForTheStaticContractPhase() {
    assertDoesNotThrow(
        () ->
            new InspectionStep(
                new StepId("bad-target"),
                new RangeSelector.ByRange("Budget", "A1:B2"),
                new WorkbookIntrospectionQuery.GetWorkbookSummary()));
  }
}
