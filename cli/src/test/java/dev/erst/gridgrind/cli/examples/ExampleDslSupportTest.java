package dev.erst.gridgrind.cli.examples;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.TableSelector;
import org.junit.jupiter.api.Test;

/** Focused tests for package-private example DSL helpers. */
class ExampleDslSupportTest {
  private static final String BUDGET_TABLE = "BudgetTable";
  private static final String BUDGET_SHEET = "Budget";
  private static final String REGIONAL_TOTALS = "RegionalTotals";
  private static final String RANGE_REPORT = "RangeReport";
  private static final String PIVOT_SOURCE = "PivotSource";
  private static final String REGION = "Region";
  private static final String AMOUNT = "Amount";

  @Test
  void tableHelperBuildsSheetOwnedTableSelectors() {
    TableSelector.ByNameOnSheet selector = ExampleSelectors.table(BUDGET_TABLE, BUDGET_SHEET);

    assertEquals(BUDGET_TABLE, selector.name());
    assertEquals(BUDGET_SHEET, selector.sheetName());
  }

  @Test
  void defaultExecutionPlanBuildsIdentifiedPlansWithRequestedPersistence() {
    WorkbookPlan plan =
        ExampleWorkbookPlans.defaultExecutionPlan(
            "budget-plan",
            new WorkbookPlan.WorkbookSource.New(),
            ExampleWorkbookPlans.saveAs("budget.xlsx"));

    assertEquals("budget-plan", plan.planId().orElseThrow());
    assertInstanceOf(WorkbookPlan.WorkbookPersistence.SaveAs.class, plan.persistence());
  }

  @Test
  void clusteredColumnComparisonChartBuildsTwoSeriesChartInputs() {
    ChartInput chart =
        ExampleChartInputs.clusteredColumnComparisonChart(
            "OpsChart",
            ExampleDrawingAnchors.anchor(1, 2, 3, 4),
            "Roadmap",
            "ChartCategories",
            "Plan",
            "Ops!$B$2:$B$4",
            "Actual",
            "ChartActual");

    assertEquals("OpsChart", chart.name());
    assertEquals(1, chart.plots().size());
    assertEquals(
        2,
        ((dev.erst.gridgrind.contract.dto.ChartPlotInput.Bar) chart.plots().getFirst())
            .series()
            .size());
  }

  @Test
  void regionalTotalsPivotFromTableBuildsExpectedPivotAxes() {
    PivotTableInput pivot =
        ExamplePivotInputs.regionalTotalsPivotFromTable(
            REGIONAL_TOTALS, RANGE_REPORT, PIVOT_SOURCE);

    assertEquals(REGIONAL_TOTALS, pivot.name());
    assertInstanceOf(PivotTableInput.Source.Table.class, pivot.source());
    assertEquals(java.util.List.of(REGION), pivot.rowLabels());
    assertTrue(
        pivot.dataFields().stream().anyMatch(field -> AMOUNT.equals(field.sourceColumnName())));
  }
}
