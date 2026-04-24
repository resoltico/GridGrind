package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.excel.WorkbookReadResult;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Evaluates first-class assertion steps by adapting them onto canonical inspection reads. */
final class AssertionExecutor {
  private final WorkbookReadExecutor readExecutor;
  private final SemanticSelectorResolver selectorResolver;

  AssertionExecutor(WorkbookReadExecutor readExecutor, SemanticSelectorResolver selectorResolver) {
    this.readExecutor = Objects.requireNonNull(readExecutor, "readExecutor must not be null");
    this.selectorResolver =
        Objects.requireNonNull(selectorResolver, "selectorResolver must not be null");
  }

  AssertionResult execute(
      AssertionStep step, ExcelWorkbook workbook, WorkbookLocation workbookLocation)
      throws AssertionFailedException {
    Objects.requireNonNull(step, "step must not be null");
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    Evaluation evaluation =
        evaluate(step.stepId(), step.target(), step.assertion(), workbook, workbookLocation);
    if (!evaluation.passed()) {
      throw new AssertionFailedException(
          evaluation.message(),
          new AssertionFailure(
              step.stepId(),
              step.assertion().assertionType(),
              step.target(),
              step.assertion(),
              evaluation.observations()));
    }
    return new AssertionResult(step.stepId(), step.assertion().assertionType());
  }

  private Evaluation evaluate(
      String stepId,
      Selector target,
      Assertion assertion,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    return switch (assertion) {
      case Assertion.NamedRangePresent namedRangePresent ->
          evaluateEntityPresence(
              stepId, target, namedRangePresent.assertionType(), true, workbook, workbookLocation);
      case Assertion.NamedRangeAbsent namedRangeAbsent ->
          evaluateEntityPresence(
              stepId, target, namedRangeAbsent.assertionType(), false, workbook, workbookLocation);
      case Assertion.TablePresent tablePresent ->
          evaluateEntityPresence(
              stepId, target, tablePresent.assertionType(), true, workbook, workbookLocation);
      case Assertion.TableAbsent tableAbsent ->
          evaluateEntityPresence(
              stepId, target, tableAbsent.assertionType(), false, workbook, workbookLocation);
      case Assertion.PivotTablePresent pivotTablePresent ->
          evaluateEntityPresence(
              stepId, target, pivotTablePresent.assertionType(), true, workbook, workbookLocation);
      case Assertion.PivotTableAbsent pivotTableAbsent ->
          evaluateEntityPresence(
              stepId, target, pivotTableAbsent.assertionType(), false, workbook, workbookLocation);
      case Assertion.ChartPresent chartPresent ->
          evaluateEntityPresence(
              stepId, target, chartPresent.assertionType(), true, workbook, workbookLocation);
      case Assertion.ChartAbsent chartAbsent ->
          evaluateEntityPresence(
              stepId, target, chartAbsent.assertionType(), false, workbook, workbookLocation);
      case Assertion.CellValue cellValue ->
          evaluateCellValue(stepId, target, cellValue.expectedValue(), workbook, workbookLocation);
      case Assertion.DisplayValue displayValue ->
          evaluateDisplayValue(
              stepId, target, displayValue.displayValue(), workbook, workbookLocation);
      case Assertion.FormulaText formulaText ->
          evaluateFormulaText(stepId, target, formulaText.formula(), workbook, workbookLocation);
      case Assertion.CellStyle cellStyle ->
          evaluateCellStyle(stepId, target, cellStyle.style(), workbook, workbookLocation);
      case Assertion.WorkbookProtectionFacts workbookProtectionFacts ->
          evaluateWorkbookProtection(
              stepId, target, workbookProtectionFacts.protection(), workbook, workbookLocation);
      case Assertion.SheetStructureFacts sheetStructureFacts ->
          evaluateSheetStructure(
              stepId, target, sheetStructureFacts.sheet(), workbook, workbookLocation);
      case Assertion.NamedRangeFacts namedRangeFacts ->
          evaluateNamedRangeFacts(
              stepId, target, namedRangeFacts.namedRanges(), workbook, workbookLocation);
      case Assertion.TableFacts tableFacts ->
          evaluateTableFacts(stepId, target, tableFacts.tables(), workbook, workbookLocation);
      case Assertion.PivotTableFacts pivotTableFacts ->
          evaluatePivotFacts(
              stepId, target, pivotTableFacts.pivotTables(), workbook, workbookLocation);
      case Assertion.ChartFacts chartFacts ->
          evaluateChartFacts(stepId, target, chartFacts.charts(), workbook, workbookLocation);
      case Assertion.AnalysisMaxSeverity analysisMaxSeverity ->
          evaluateAnalysisMaxSeverity(
              stepId,
              target,
              analysisMaxSeverity.query(),
              analysisMaxSeverity.maximumSeverity(),
              workbook,
              workbookLocation);
      case Assertion.AnalysisFindingPresent analysisFindingPresent ->
          evaluateFindingPresence(
              stepId,
              target,
              analysisFindingPresent.query(),
              analysisFindingPresent.code(),
              analysisFindingPresent.severity(),
              analysisFindingPresent.messageContains(),
              true,
              workbook,
              workbookLocation);
      case Assertion.AnalysisFindingAbsent analysisFindingAbsent ->
          evaluateFindingPresence(
              stepId,
              target,
              analysisFindingAbsent.query(),
              analysisFindingAbsent.code(),
              analysisFindingAbsent.severity(),
              analysisFindingAbsent.messageContains(),
              false,
              workbook,
              workbookLocation);
      case Assertion.AllOf allOf ->
          evaluateAllOf(stepId, target, allOf.assertions(), workbook, workbookLocation);
      case Assertion.AnyOf anyOf ->
          evaluateAnyOf(stepId, target, anyOf.assertions(), workbook, workbookLocation);
      case Assertion.Not not ->
          evaluateNot(stepId, target, not.assertion(), workbook, workbookLocation);
    };
  }

  private Evaluation evaluateEntityPresence(
      String stepId,
      Selector target,
      String assertionType,
      boolean shouldExist,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    List<InspectionResult> observations =
        List.of(presenceObservation(stepId, target, workbook, workbookLocation));
    int count = observedCount(observations.getFirst());
    boolean matchedExpectation = shouldExist ? count > 0 : count == 0;
    return matchedExpectation
        ? Evaluation.pass(observations)
        : Evaluation.fail(
            observations,
            shouldExist
                ? assertionType + " observed no matching workbook entities"
                : assertionType + " observed " + count + " matching workbook entities");
  }

  private Evaluation evaluateCellValue(
      String stepId,
      Selector target,
      ExpectedCellValue expectedValue,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.CellsResult cellsResult =
        (InspectionResult.CellsResult)
            executeObservation(
                stepId, target, new InspectionQuery.GetCells(), workbook, workbookLocation);
    if (cellsResult.cells().isEmpty()) {
      return Evaluation.fail(
          List.of(cellsResult), "EXPECT_CELL_VALUE resolved no matching cells to compare");
    }
    List<String> mismatches =
        cellsResult.cells().stream()
            .filter(cell -> !matchesCellValue(cell, expectedValue))
            .map(GridGrindResponse.CellReport::address)
            .toList();
    return mismatches.isEmpty()
        ? Evaluation.pass(List.of(cellsResult))
        : Evaluation.fail(
            List.of(cellsResult),
            "EXPECT_CELL_VALUE mismatched effective values at " + String.join(", ", mismatches));
  }

  private Evaluation evaluateDisplayValue(
      String stepId,
      Selector target,
      String expectedDisplayValue,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.CellsResult cellsResult =
        (InspectionResult.CellsResult)
            executeObservation(
                stepId, target, new InspectionQuery.GetCells(), workbook, workbookLocation);
    if (cellsResult.cells().isEmpty()) {
      return Evaluation.fail(
          List.of(cellsResult), "EXPECT_DISPLAY_VALUE resolved no matching cells to compare");
    }
    List<String> mismatches =
        cellsResult.cells().stream()
            .filter(cell -> !cell.displayValue().equals(expectedDisplayValue))
            .map(GridGrindResponse.CellReport::address)
            .toList();
    return mismatches.isEmpty()
        ? Evaluation.pass(List.of(cellsResult))
        : Evaluation.fail(
            List.of(cellsResult),
            "EXPECT_DISPLAY_VALUE mismatched formatted values at " + String.join(", ", mismatches));
  }

  private Evaluation evaluateFormulaText(
      String stepId,
      Selector target,
      String expectedFormula,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.CellsResult cellsResult =
        (InspectionResult.CellsResult)
            executeObservation(
                stepId, target, new InspectionQuery.GetCells(), workbook, workbookLocation);
    if (cellsResult.cells().isEmpty()) {
      return Evaluation.fail(
          List.of(cellsResult), "EXPECT_FORMULA_TEXT resolved no matching cells to compare");
    }
    List<String> mismatches =
        cellsResult.cells().stream()
            .filter(
                cell ->
                    !(cell instanceof GridGrindResponse.CellReport.FormulaReport formulaReport)
                        || !formulaReport.formula().equals(expectedFormula))
            .map(GridGrindResponse.CellReport::address)
            .toList();
    return mismatches.isEmpty()
        ? Evaluation.pass(List.of(cellsResult))
        : Evaluation.fail(
            List.of(cellsResult),
            "EXPECT_FORMULA_TEXT mismatched formula cells at " + String.join(", ", mismatches));
  }

  private Evaluation evaluateCellStyle(
      String stepId,
      Selector target,
      GridGrindResponse.CellStyleReport expectedStyle,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.CellsResult cellsResult =
        (InspectionResult.CellsResult)
            executeObservation(
                stepId, target, new InspectionQuery.GetCells(), workbook, workbookLocation);
    if (cellsResult.cells().isEmpty()) {
      return Evaluation.fail(
          List.of(cellsResult), "EXPECT_CELL_STYLE resolved no matching cells to compare");
    }
    List<String> mismatches =
        cellsResult.cells().stream()
            .filter(cell -> !cell.style().equals(expectedStyle))
            .map(GridGrindResponse.CellReport::address)
            .toList();
    return mismatches.isEmpty()
        ? Evaluation.pass(List.of(cellsResult))
        : Evaluation.fail(
            List.of(cellsResult),
            "EXPECT_CELL_STYLE mismatched style snapshots at " + String.join(", ", mismatches));
  }

  private Evaluation evaluateWorkbookProtection(
      String stepId,
      Selector target,
      dev.erst.gridgrind.contract.dto.WorkbookProtectionReport expectedProtection,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.WorkbookProtectionResult result =
        (InspectionResult.WorkbookProtectionResult)
            executeObservation(
                stepId,
                target,
                new InspectionQuery.GetWorkbookProtection(),
                workbook,
                workbookLocation);
    return result.protection().equals(expectedProtection)
        ? Evaluation.pass(List.of(result))
        : Evaluation.fail(
            List.of(result), "EXPECT_WORKBOOK_PROTECTION observed a different protection report");
  }

  private Evaluation evaluateSheetStructure(
      String stepId,
      Selector target,
      GridGrindResponse.SheetSummaryReport expectedSheet,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.SheetSummaryResult result =
        (InspectionResult.SheetSummaryResult)
            executeObservation(
                stepId, target, new InspectionQuery.GetSheetSummary(), workbook, workbookLocation);
    return result.sheet().equals(expectedSheet)
        ? Evaluation.pass(List.of(result))
        : Evaluation.fail(
            List.of(result), "EXPECT_SHEET_STRUCTURE observed a different sheet summary report");
  }

  private Evaluation evaluateNamedRangeFacts(
      String stepId,
      Selector target,
      List<GridGrindResponse.NamedRangeReport> expectedNamedRanges,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.NamedRangesResult result =
        (InspectionResult.NamedRangesResult)
            executeObservation(
                stepId, target, new InspectionQuery.GetNamedRanges(), workbook, workbookLocation);
    return result.namedRanges().equals(expectedNamedRanges)
        ? Evaluation.pass(List.of(result))
        : Evaluation.fail(
            List.of(result), "EXPECT_NAMED_RANGE_FACTS observed different named-range reports");
  }

  private Evaluation evaluateTableFacts(
      String stepId,
      Selector target,
      List<dev.erst.gridgrind.contract.dto.TableEntryReport> expectedTables,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.TablesResult result =
        (InspectionResult.TablesResult)
            executeObservation(
                stepId, target, new InspectionQuery.GetTables(), workbook, workbookLocation);
    return result.tables().equals(expectedTables)
        ? Evaluation.pass(List.of(result))
        : Evaluation.fail(List.of(result), "EXPECT_TABLE_FACTS observed different table reports");
  }

  private Evaluation evaluatePivotFacts(
      String stepId,
      Selector target,
      List<dev.erst.gridgrind.contract.dto.PivotTableReport> expectedPivotTables,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.PivotTablesResult result =
        (InspectionResult.PivotTablesResult)
            executeObservation(
                stepId, target, new InspectionQuery.GetPivotTables(), workbook, workbookLocation);
    return result.pivotTables().equals(expectedPivotTables)
        ? Evaluation.pass(List.of(result))
        : Evaluation.fail(
            List.of(result), "EXPECT_PIVOT_TABLE_FACTS observed different pivot-table reports");
  }

  private Evaluation evaluateChartFacts(
      String stepId,
      Selector target,
      List<ChartReport> expectedCharts,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.ChartsResult result =
        chartsObservation(stepId, target, workbook, workbookLocation);
    return result.charts().equals(expectedCharts)
        ? Evaluation.pass(List.of(result))
        : Evaluation.fail(List.of(result), "EXPECT_CHART_FACTS observed different chart reports");
  }

  private Evaluation evaluateAnalysisMaxSeverity(
      String stepId,
      Selector target,
      InspectionQuery query,
      AnalysisSeverity maximumSeverity,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.Analysis result =
        (InspectionResult.Analysis)
            executeObservation(stepId, target, query, workbook, workbookLocation);
    Optional<AnalysisSeverity> observedSeverity = highestSeverity(result);
    return severityRank(observedSeverity) <= severityRank(Optional.of(maximumSeverity))
        ? Evaluation.pass(List.of(result))
        : Evaluation.fail(
            List.of(result),
            "EXPECT_ANALYSIS_MAX_SEVERITY observed highest severity "
                + observedSeverity.map(Enum::name).orElse("NONE")
                + " which exceeds "
                + maximumSeverity);
  }

  private Evaluation evaluateFindingPresence(
      String stepId,
      Selector target,
      InspectionQuery query,
      dev.erst.gridgrind.contract.dto.AnalysisFindingCode code,
      AnalysisSeverity severity,
      String messageContains,
      boolean shouldExist,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.Analysis result =
        (InspectionResult.Analysis)
            executeObservation(stepId, target, query, workbook, workbookLocation);
    boolean found =
        analysisFindings(result).stream()
            .anyMatch(finding -> matchesFinding(finding, code, severity, messageContains));
    if (found == shouldExist) {
      return Evaluation.pass(List.of(result));
    }
    String verb = shouldExist ? "missing" : "unexpectedly present";
    return Evaluation.fail(List.of(result), query.queryType() + " " + verb + " finding " + code);
  }

  private Evaluation evaluateAllOf(
      String stepId,
      Selector target,
      List<Assertion> assertions,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    List<InspectionResult> observations = new ArrayList<>();
    List<String> failures = new ArrayList<>();
    for (Assertion nestedAssertion : assertions) {
      Evaluation evaluation = evaluate(stepId, target, nestedAssertion, workbook, workbookLocation);
      observations.addAll(evaluation.observations());
      if (!evaluation.passed()) {
        failures.add(nestedAssertion.assertionType() + ": " + evaluation.message());
      }
    }
    return failures.isEmpty()
        ? Evaluation.pass(observations)
        : Evaluation.fail(observations, "ALL_OF failed for " + String.join("; ", failures));
  }

  private Evaluation evaluateAnyOf(
      String stepId,
      Selector target,
      List<Assertion> assertions,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    List<InspectionResult> observations = new ArrayList<>();
    List<String> failures = new ArrayList<>();
    for (Assertion nestedAssertion : assertions) {
      Evaluation evaluation = evaluate(stepId, target, nestedAssertion, workbook, workbookLocation);
      observations.addAll(evaluation.observations());
      if (evaluation.passed()) {
        return Evaluation.pass(observations);
      }
      failures.add(nestedAssertion.assertionType() + ": " + evaluation.message());
    }
    return Evaluation.fail(observations, "ANY_OF failed for " + String.join("; ", failures));
  }

  private Evaluation evaluateNot(
      String stepId,
      Selector target,
      Assertion nestedAssertion,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    Evaluation evaluation = evaluate(stepId, target, nestedAssertion, workbook, workbookLocation);
    return evaluation.passed()
        ? Evaluation.fail(
            evaluation.observations(),
            "NOT failed because nested assertion " + nestedAssertion.assertionType() + " passed")
        : Evaluation.pass(evaluation.observations());
  }

  private InspectionResult executeObservation(
      String stepId,
      Selector target,
      InspectionQuery query,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    SemanticSelectorResolver.ResolvedInspectionTarget resolvedTarget =
        selectorResolver.resolveInspectionTarget(stepId, workbook, target, query);
    if (resolvedTarget.isShortCircuit()) {
      return resolvedTarget.shortCircuitResult();
    }
    WorkbookReadResult result =
        readExecutor
            .apply(
                workbook,
                workbookLocation,
                InspectionCommandConverter.toReadCommand(stepId, resolvedTarget.selector(), query))
            .getFirst();
    return InspectionResultConverter.toReadResult(result);
  }

  InspectionResult presenceObservation(
      String stepId, Selector target, ExcelWorkbook workbook, WorkbookLocation workbookLocation) {
    try {
      return switch (target) {
        case NamedRangeSelector selector ->
            executeObservation(
                stepId, selector, new InspectionQuery.GetNamedRanges(), workbook, workbookLocation);
        case TableSelector selector ->
            executeObservation(
                stepId, selector, new InspectionQuery.GetTables(), workbook, workbookLocation);
        case PivotTableSelector selector ->
            executeObservation(
                stepId, selector, new InspectionQuery.GetPivotTables(), workbook, workbookLocation);
        case ChartSelector _ -> chartsObservation(stepId, target, workbook, workbookLocation);
        default ->
            throw new IllegalArgumentException(
                "Unsupported presence assertion target: " + target.getClass().getSimpleName());
      };
    } catch (NamedRangeNotFoundException | SheetNotFoundException ignored) {
      return zeroMatchPresenceObservation(stepId, target);
    }
  }

  InspectionResult.ChartsResult chartsObservation(
      String stepId, Selector target, ExcelWorkbook workbook, WorkbookLocation workbookLocation) {
    String sheetName =
        switch (target) {
          case ChartSelector.AllOnSheet selector -> selector.sheetName();
          case ChartSelector.ByName selector -> selector.sheetName();
          default -> throw new IllegalArgumentException("Unsupported chart assertion target");
        };
    InspectionResult.ChartsResult allCharts =
        (InspectionResult.ChartsResult)
            executeObservation(
                stepId,
                new SheetSelector.ByName(sheetName),
                new InspectionQuery.GetCharts(),
                workbook,
                workbookLocation);
    if (target instanceof ChartSelector.ByName selector) {
      return new InspectionResult.ChartsResult(
          stepId,
          allCharts.sheetName(),
          allCharts.charts().stream()
              .filter(chart -> chart.name().equals(selector.chartName()))
              .toList());
    }
    return allCharts;
  }

  static InspectionResult zeroMatchPresenceObservation(String stepId, Selector target) {
    return switch (target) {
      case NamedRangeSelector _ -> new InspectionResult.NamedRangesResult(stepId, List.of());
      case TableSelector _ -> new InspectionResult.TablesResult(stepId, List.of());
      case PivotTableSelector _ -> new InspectionResult.PivotTablesResult(stepId, List.of());
      case ChartSelector.AllOnSheet selector ->
          new InspectionResult.ChartsResult(stepId, selector.sheetName(), List.of());
      case ChartSelector.ByName selector ->
          new InspectionResult.ChartsResult(stepId, selector.sheetName(), List.of());
      default ->
          throw new IllegalArgumentException(
              "Unsupported presence assertion target: " + target.getClass().getSimpleName());
    };
  }

  static int observedCount(InspectionResult observation) {
    return switch (observation) {
      case InspectionResult.NamedRangesResult result -> result.namedRanges().size();
      case InspectionResult.TablesResult result -> result.tables().size();
      case InspectionResult.PivotTablesResult result -> result.pivotTables().size();
      case InspectionResult.ChartsResult result -> result.charts().size();
      default ->
          throw new IllegalArgumentException(
              "Unsupported presence observation result: " + observation.getClass().getSimpleName());
    };
  }

  static boolean matchesCellValue(
      GridGrindResponse.CellReport cell, ExpectedCellValue expectedValue) {
    if (cell instanceof GridGrindResponse.CellReport.FormulaReport formulaReport) {
      return matchesCellValue(formulaReport.evaluation(), expectedValue);
    }
    return switch (expectedValue) {
      case ExpectedCellValue.Blank _ -> cell instanceof GridGrindResponse.CellReport.BlankReport;
      case ExpectedCellValue.Text expectedText ->
          cell instanceof GridGrindResponse.CellReport.TextReport textReport
              && textReport.stringValue().equals(expectedText.text());
      case ExpectedCellValue.NumericValue expectedNumber ->
          cell instanceof GridGrindResponse.CellReport.NumberReport numberReport
              && Double.compare(numberReport.numberValue(), expectedNumber.number()) == 0;
      case ExpectedCellValue.BooleanValue expectedBoolean ->
          cell instanceof GridGrindResponse.CellReport.BooleanReport booleanReport
              && booleanReport.booleanValue().equals(expectedBoolean.value());
      case ExpectedCellValue.ErrorValue expectedError ->
          cell instanceof GridGrindResponse.CellReport.ErrorReport errorReport
              && errorReport.errorValue().equals(expectedError.error());
    };
  }

  static Optional<AnalysisSeverity> highestSeverity(InspectionResult.Analysis result) {
    GridGrindResponse.AnalysisSummaryReport summary = analysisSummary(result);
    if (summary.errorCount() > 0) {
      return Optional.of(AnalysisSeverity.ERROR);
    }
    if (summary.warningCount() > 0) {
      return Optional.of(AnalysisSeverity.WARNING);
    }
    if (summary.infoCount() > 0) {
      return Optional.of(AnalysisSeverity.INFO);
    }
    return Optional.empty();
  }

  static int severityRank(Optional<AnalysisSeverity> severity) {
    Objects.requireNonNull(severity, "severity must not be null");
    return severity.map(AssertionExecutor::severityRank).orElse(-1);
  }

  private static int severityRank(AnalysisSeverity severity) {
    return switch (Objects.requireNonNull(severity, "severity must not be null")) {
      case INFO -> 0;
      case WARNING -> 1;
      case ERROR -> 2;
    };
  }

  static GridGrindResponse.AnalysisSummaryReport analysisSummary(InspectionResult.Analysis result) {
    return switch (result) {
      case InspectionResult.FormulaHealthResult formulaHealthResult ->
          formulaHealthResult.analysis().summary();
      case InspectionResult.DataValidationHealthResult dataValidationHealthResult ->
          dataValidationHealthResult.analysis().summary();
      case InspectionResult.ConditionalFormattingHealthResult conditionalFormattingHealthResult ->
          conditionalFormattingHealthResult.analysis().summary();
      case InspectionResult.AutofilterHealthResult autofilterHealthResult ->
          autofilterHealthResult.analysis().summary();
      case InspectionResult.TableHealthResult tableHealthResult ->
          tableHealthResult.analysis().summary();
      case InspectionResult.PivotTableHealthResult pivotTableHealthResult ->
          pivotTableHealthResult.analysis().summary();
      case InspectionResult.HyperlinkHealthResult hyperlinkHealthResult ->
          hyperlinkHealthResult.analysis().summary();
      case InspectionResult.NamedRangeHealthResult namedRangeHealthResult ->
          namedRangeHealthResult.analysis().summary();
      case InspectionResult.WorkbookFindingsResult workbookFindingsResult ->
          workbookFindingsResult.analysis().summary();
    };
  }

  static List<GridGrindResponse.AnalysisFindingReport> analysisFindings(
      InspectionResult.Analysis result) {
    return switch (result) {
      case InspectionResult.FormulaHealthResult formulaHealthResult ->
          formulaHealthResult.analysis().findings();
      case InspectionResult.DataValidationHealthResult dataValidationHealthResult ->
          dataValidationHealthResult.analysis().findings();
      case InspectionResult.ConditionalFormattingHealthResult conditionalFormattingHealthResult ->
          conditionalFormattingHealthResult.analysis().findings();
      case InspectionResult.AutofilterHealthResult autofilterHealthResult ->
          autofilterHealthResult.analysis().findings();
      case InspectionResult.TableHealthResult tableHealthResult ->
          tableHealthResult.analysis().findings();
      case InspectionResult.PivotTableHealthResult pivotTableHealthResult ->
          pivotTableHealthResult.analysis().findings();
      case InspectionResult.HyperlinkHealthResult hyperlinkHealthResult ->
          hyperlinkHealthResult.analysis().findings();
      case InspectionResult.NamedRangeHealthResult namedRangeHealthResult ->
          namedRangeHealthResult.analysis().findings();
      case InspectionResult.WorkbookFindingsResult workbookFindingsResult ->
          workbookFindingsResult.analysis().findings();
    };
  }

  static boolean matchesFinding(
      GridGrindResponse.AnalysisFindingReport finding,
      dev.erst.gridgrind.contract.dto.AnalysisFindingCode code,
      AnalysisSeverity severity,
      String messageContains) {
    if (finding.code() != code) {
      return false;
    }
    if (severity != null && finding.severity() != severity) {
      return false;
    }
    return messageContains == null || finding.message().contains(messageContains);
  }

  private record Evaluation(boolean passed, List<InspectionResult> observations, String message) {
    private Evaluation {
      observations = List.copyOf(observations);
      Objects.requireNonNull(message, "message must not be null");
    }

    static Evaluation pass(List<InspectionResult> observations) {
      return new Evaluation(true, observations, "passed");
    }

    static Evaluation fail(List<InspectionResult> observations, String message) {
      return new Evaluation(false, observations, message);
    }
  }
}
