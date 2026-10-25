package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.util.List;
import java.util.Objects;

/** Evaluates fact-report assertion families against canonical workbook inspections. */
final class AssertionFactEvaluator {
  private final AssertionObservationExecutor observations;

  AssertionFactEvaluator(AssertionObservationExecutor observations) {
    this.observations = Objects.requireNonNull(observations, "observations must not be null");
  }

  AssertionEvaluation evaluateWorkbookProtection(
      StepId stepId,
      Selector target,
      dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport expectedProtection,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    WorkbookInspectionResult.WorkbookProtectionResult result =
        (WorkbookInspectionResult.WorkbookProtectionResult)
            observations.executeObservation(
                stepId,
                target,
                new WorkbookIntrospectionQuery.GetWorkbookProtection(),
                workbook,
                workbookLocation);
    return result.protection().equals(expectedProtection)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_WORKBOOK_PROTECTION observed a different protection report");
  }

  AssertionEvaluation evaluateSheetStructure(
      StepId stepId,
      Selector target,
      SheetSummaryReport expectedSheet,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    SheetInspectionResult.SheetSummaryResult result =
        (SheetInspectionResult.SheetSummaryResult)
            observations.executeObservation(
                stepId,
                target,
                new SheetIntrospectionQuery.GetSheetSummary(),
                workbook,
                workbookLocation);
    return result.sheet().equals(expectedSheet)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_SHEET_STRUCTURE observed a different sheet summary report");
  }

  AssertionEvaluation evaluateNamedRangeFacts(
      StepId stepId,
      Selector target,
      List<NamedRangeReport> expectedNamedRanges,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    WorkbookInspectionResult.NamedRangesResult result =
        (WorkbookInspectionResult.NamedRangesResult)
            observations.executeObservation(
                stepId,
                target,
                new WorkbookIntrospectionQuery.GetNamedRanges(),
                workbook,
                workbookLocation);
    return result.namedRanges().equals(expectedNamedRanges)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_NAMED_RANGE_FACTS observed different named-range reports");
  }

  AssertionEvaluation evaluateTableFacts(
      StepId stepId,
      Selector target,
      List<dev.erst.gridgrind.core.model.table.TableEntryReport> expectedTables,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    WorkbookAssetInspectionResult.TablesResult result =
        (WorkbookAssetInspectionResult.TablesResult)
            observations.executeObservation(
                stepId,
                target,
                new WorkbookAssetIntrospectionQuery.GetTables(),
                workbook,
                workbookLocation);
    return result.tables().equals(expectedTables)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_TABLE_FACTS observed different table reports");
  }

  AssertionEvaluation evaluatePivotFacts(
      StepId stepId,
      Selector target,
      List<dev.erst.gridgrind.core.model.pivot.PivotTableReport> expectedPivotTables,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    WorkbookAssetInspectionResult.PivotTablesResult result =
        (WorkbookAssetInspectionResult.PivotTablesResult)
            observations.executeObservation(
                stepId,
                target,
                new WorkbookAssetIntrospectionQuery.GetPivotTables(),
                workbook,
                workbookLocation);
    return result.pivotTables().equals(expectedPivotTables)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_PIVOT_TABLE_FACTS observed different pivot-table reports");
  }

  AssertionEvaluation evaluateChartFacts(
      StepId stepId,
      Selector target,
      List<ChartReport> expectedCharts,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    WorkbookAssetInspectionResult.ChartsResult result =
        (WorkbookAssetInspectionResult.ChartsResult)
            observations.executeObservation(
                stepId,
                target,
                new WorkbookAssetIntrospectionQuery.GetCharts(),
                workbook,
                workbookLocation);
    return result.charts().equals(expectedCharts)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_CHART_FACTS observed different chart reports");
  }
}
