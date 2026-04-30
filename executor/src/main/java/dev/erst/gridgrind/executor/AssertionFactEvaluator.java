package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.Selector;
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
      String stepId,
      Selector target,
      dev.erst.gridgrind.contract.dto.WorkbookProtectionReport expectedProtection,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.WorkbookProtectionResult result =
        (InspectionResult.WorkbookProtectionResult)
            observations.executeObservation(
                stepId,
                target,
                new InspectionQuery.GetWorkbookProtection(),
                workbook,
                workbookLocation);
    return result.protection().equals(expectedProtection)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_WORKBOOK_PROTECTION observed a different protection report");
  }

  AssertionEvaluation evaluateSheetStructure(
      String stepId,
      Selector target,
      GridGrindWorkbookSurfaceReports.SheetSummaryReport expectedSheet,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.SheetSummaryResult result =
        (InspectionResult.SheetSummaryResult)
            observations.executeObservation(
                stepId, target, new InspectionQuery.GetSheetSummary(), workbook, workbookLocation);
    return result.sheet().equals(expectedSheet)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_SHEET_STRUCTURE observed a different sheet summary report");
  }

  AssertionEvaluation evaluateNamedRangeFacts(
      String stepId,
      Selector target,
      List<GridGrindWorkbookSurfaceReports.NamedRangeReport> expectedNamedRanges,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.NamedRangesResult result =
        (InspectionResult.NamedRangesResult)
            observations.executeObservation(
                stepId, target, new InspectionQuery.GetNamedRanges(), workbook, workbookLocation);
    return result.namedRanges().equals(expectedNamedRanges)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_NAMED_RANGE_FACTS observed different named-range reports");
  }

  AssertionEvaluation evaluateTableFacts(
      String stepId,
      Selector target,
      List<dev.erst.gridgrind.contract.dto.TableEntryReport> expectedTables,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.TablesResult result =
        (InspectionResult.TablesResult)
            observations.executeObservation(
                stepId, target, new InspectionQuery.GetTables(), workbook, workbookLocation);
    return result.tables().equals(expectedTables)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_TABLE_FACTS observed different table reports");
  }

  AssertionEvaluation evaluatePivotFacts(
      String stepId,
      Selector target,
      List<dev.erst.gridgrind.contract.dto.PivotTableReport> expectedPivotTables,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.PivotTablesResult result =
        (InspectionResult.PivotTablesResult)
            observations.executeObservation(
                stepId, target, new InspectionQuery.GetPivotTables(), workbook, workbookLocation);
    return result.pivotTables().equals(expectedPivotTables)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_PIVOT_TABLE_FACTS observed different pivot-table reports");
  }

  AssertionEvaluation evaluateChartFacts(
      String stepId,
      Selector target,
      List<ChartReport> expectedCharts,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.ChartsResult result =
        observations.chartsObservation(stepId, target, workbook, workbookLocation);
    return result.charts().equals(expectedCharts)
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result), "EXPECT_CHART_FACTS observed different chart reports");
  }
}
