package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AnalysisAssertion;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.assertion.CompositeAssertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.assertion.PresenceAssertion;
import dev.erst.gridgrind.contract.assertion.WorkbookFactAssertion;
import dev.erst.gridgrind.contract.dto.AnalysisFindingReport;
import dev.erst.gridgrind.contract.dto.AnalysisSummaryReport;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookAnalysisResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Evaluates first-class assertion steps by adapting them onto canonical inspection reads. */
final class AssertionExecutor {
  private final AssertionObservationExecutor observations;
  private final AssertionValueEvaluator valueEvaluator;
  private final AssertionFactEvaluator factEvaluator;
  private final AssertionAnalysisEvaluator analysisEvaluator;

  AssertionExecutor(
      WorkbookExecutionEngine workbookEngine, SemanticSelectorResolver selectorResolver) {
    this.observations = new AssertionObservationExecutor(workbookEngine, selectorResolver);
    this.valueEvaluator = new AssertionValueEvaluator(observations);
    this.factEvaluator = new AssertionFactEvaluator(observations);
    this.analysisEvaluator = new AssertionAnalysisEvaluator(observations);
  }

  AssertionResult execute(
      AssertionStep step, ExcelWorkbook workbook, WorkbookLocation workbookLocation)
      throws AssertionFailedException {
    Objects.requireNonNull(step, "step must not be null");
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    AssertionEvaluation evaluation =
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

  private AssertionEvaluation evaluate(
      String stepId,
      Selector target,
      Assertion assertion,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    return switch (assertion) {
      case PresenceAssertion.NamedRangePresent namedRangePresent ->
          valueEvaluator.evaluateEntityPresence(
              stepId, target, namedRangePresent.assertionType(), true, workbook, workbookLocation);
      case PresenceAssertion.NamedRangeAbsent namedRangeAbsent ->
          valueEvaluator.evaluateEntityPresence(
              stepId, target, namedRangeAbsent.assertionType(), false, workbook, workbookLocation);
      case PresenceAssertion.TablePresent tablePresent ->
          valueEvaluator.evaluateEntityPresence(
              stepId, target, tablePresent.assertionType(), true, workbook, workbookLocation);
      case PresenceAssertion.TableAbsent tableAbsent ->
          valueEvaluator.evaluateEntityPresence(
              stepId, target, tableAbsent.assertionType(), false, workbook, workbookLocation);
      case PresenceAssertion.PivotTablePresent pivotTablePresent ->
          valueEvaluator.evaluateEntityPresence(
              stepId, target, pivotTablePresent.assertionType(), true, workbook, workbookLocation);
      case PresenceAssertion.PivotTableAbsent pivotTableAbsent ->
          valueEvaluator.evaluateEntityPresence(
              stepId, target, pivotTableAbsent.assertionType(), false, workbook, workbookLocation);
      case PresenceAssertion.ChartPresent chartPresent ->
          valueEvaluator.evaluateEntityPresence(
              stepId, target, chartPresent.assertionType(), true, workbook, workbookLocation);
      case PresenceAssertion.ChartAbsent chartAbsent ->
          valueEvaluator.evaluateEntityPresence(
              stepId, target, chartAbsent.assertionType(), false, workbook, workbookLocation);
      case CellAssertion.CellValue cellValue ->
          valueEvaluator.evaluateCellValue(
              stepId, target, cellValue.expectedValue(), workbook, workbookLocation);
      case CellAssertion.DisplayValue displayValue ->
          valueEvaluator.evaluateDisplayValue(
              stepId, target, displayValue.displayValue(), workbook, workbookLocation);
      case CellAssertion.FormulaText formulaText ->
          valueEvaluator.evaluateFormulaText(
              stepId, target, formulaText.formula(), workbook, workbookLocation);
      case CellAssertion.CellStyle cellStyle ->
          valueEvaluator.evaluateCellStyle(
              stepId, target, cellStyle.style(), workbook, workbookLocation);
      case WorkbookFactAssertion.WorkbookProtectionFacts workbookProtectionFacts ->
          factEvaluator.evaluateWorkbookProtection(
              stepId, target, workbookProtectionFacts.protection(), workbook, workbookLocation);
      case WorkbookFactAssertion.SheetStructureFacts sheetStructureFacts ->
          factEvaluator.evaluateSheetStructure(
              stepId, target, sheetStructureFacts.sheet(), workbook, workbookLocation);
      case WorkbookFactAssertion.NamedRangeFacts namedRangeFacts ->
          factEvaluator.evaluateNamedRangeFacts(
              stepId, target, namedRangeFacts.namedRanges(), workbook, workbookLocation);
      case WorkbookFactAssertion.TableFacts tableFacts ->
          factEvaluator.evaluateTableFacts(
              stepId, target, tableFacts.tables(), workbook, workbookLocation);
      case WorkbookFactAssertion.PivotTableFacts pivotTableFacts ->
          factEvaluator.evaluatePivotFacts(
              stepId, target, pivotTableFacts.pivotTables(), workbook, workbookLocation);
      case WorkbookFactAssertion.ChartFacts chartFacts ->
          factEvaluator.evaluateChartFacts(
              stepId, target, chartFacts.charts(), workbook, workbookLocation);
      case AnalysisAssertion.AnalysisMaxSeverity analysisMaxSeverity ->
          analysisEvaluator.evaluateAnalysisMaxSeverity(
              stepId,
              target,
              analysisMaxSeverity.query(),
              analysisMaxSeverity.maximumSeverity(),
              workbook,
              workbookLocation);
      case AnalysisAssertion.AnalysisFindingPresent analysisFindingPresent ->
          analysisEvaluator.evaluateFindingPresence(
              stepId,
              target,
              analysisFindingPresent.query(),
              analysisFindingPresent.code(),
              analysisFindingPresent.severity(),
              analysisFindingPresent.messageContains(),
              true,
              workbook,
              workbookLocation);
      case AnalysisAssertion.AnalysisFindingAbsent analysisFindingAbsent ->
          analysisEvaluator.evaluateFindingPresence(
              stepId,
              target,
              analysisFindingAbsent.query(),
              analysisFindingAbsent.code(),
              analysisFindingAbsent.severity(),
              analysisFindingAbsent.messageContains(),
              false,
              workbook,
              workbookLocation);
      case CompositeAssertion.AllOf allOf ->
          evaluateAllOf(stepId, target, allOf.assertions(), workbook, workbookLocation);
      case CompositeAssertion.AnyOf anyOf ->
          evaluateAnyOf(stepId, target, anyOf.assertions(), workbook, workbookLocation);
      case CompositeAssertion.Not not ->
          evaluateNot(stepId, target, not.assertion(), workbook, workbookLocation);
    };
  }

  private AssertionEvaluation evaluateAllOf(
      String stepId,
      Selector target,
      List<Assertion> assertions,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    List<InspectionResult> observations = new ArrayList<>();
    List<String> failures = new ArrayList<>();
    for (Assertion nestedAssertion : assertions) {
      AssertionEvaluation evaluation =
          evaluate(stepId, target, nestedAssertion, workbook, workbookLocation);
      observations.addAll(evaluation.observations());
      if (!evaluation.passed()) {
        failures.add(nestedAssertion.assertionType() + ": " + evaluation.message());
      }
    }
    return failures.isEmpty()
        ? AssertionEvaluation.pass(observations)
        : AssertionEvaluation.fail(
            observations, "ALL_OF failed for " + String.join("; ", failures));
  }

  private AssertionEvaluation evaluateAnyOf(
      String stepId,
      Selector target,
      List<Assertion> assertions,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    List<InspectionResult> observations = new ArrayList<>();
    List<String> failures = new ArrayList<>();
    for (Assertion nestedAssertion : assertions) {
      AssertionEvaluation evaluation =
          evaluate(stepId, target, nestedAssertion, workbook, workbookLocation);
      observations.addAll(evaluation.observations());
      if (evaluation.passed()) {
        return AssertionEvaluation.pass(observations);
      }
      failures.add(nestedAssertion.assertionType() + ": " + evaluation.message());
    }
    return AssertionEvaluation.fail(
        observations, "ANY_OF failed for " + String.join("; ", failures));
  }

  private AssertionEvaluation evaluateNot(
      String stepId,
      Selector target,
      Assertion nestedAssertion,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    AssertionEvaluation evaluation =
        evaluate(stepId, target, nestedAssertion, workbook, workbookLocation);
    return evaluation.passed()
        ? AssertionEvaluation.fail(
            evaluation.observations(),
            "NOT failed because nested assertion " + nestedAssertion.assertionType() + " passed")
        : AssertionEvaluation.pass(evaluation.observations());
  }

  InspectionResult presenceObservation(
      String stepId, Selector target, ExcelWorkbook workbook, WorkbookLocation workbookLocation) {
    return observations.presenceObservation(stepId, target, workbook, workbookLocation);
  }

  WorkbookAssetInspectionResult.ChartsResult chartsObservation(
      String stepId, Selector target, ExcelWorkbook workbook, WorkbookLocation workbookLocation) {
    return (WorkbookAssetInspectionResult.ChartsResult)
        observations.executeObservation(
            stepId,
            target,
            new WorkbookAssetIntrospectionQuery.GetCharts(),
            workbook,
            workbookLocation);
  }

  static InspectionResult zeroMatchPresenceObservation(String stepId, Selector target) {
    return AssertionObservationExecutor.zeroMatchPresenceObservation(stepId, target);
  }

  static int observedCount(InspectionResult observation) {
    return AssertionObservationExecutor.observedCount(observation);
  }

  static boolean matchesCellValue(
      dev.erst.gridgrind.contract.dto.CellReport cell, ExpectedCellValue expectedValue) {
    return AssertionValueEvaluator.matchesCellValue(cell, expectedValue);
  }

  static java.util.Optional<AnalysisSeverity> highestSeverity(WorkbookAnalysisResult result) {
    return AssertionAnalysisEvaluator.highestSeverity(result);
  }

  static int severityRank(java.util.Optional<AnalysisSeverity> severity) {
    return AssertionAnalysisEvaluator.severityRank(severity);
  }

  static AnalysisSummaryReport analysisSummary(WorkbookAnalysisResult result) {
    return AssertionAnalysisEvaluator.analysisSummary(result);
  }

  static List<AnalysisFindingReport> analysisFindings(WorkbookAnalysisResult result) {
    return AssertionAnalysisEvaluator.analysisFindings(result);
  }

  static boolean matchesFinding(
      AnalysisFindingReport finding,
      AnalysisFindingCode code,
      java.util.Optional<AnalysisSeverity> severity,
      java.util.Optional<String> messageContains) {
    return AssertionAnalysisEvaluator.matchesFinding(finding, code, severity, messageContains);
  }
}
