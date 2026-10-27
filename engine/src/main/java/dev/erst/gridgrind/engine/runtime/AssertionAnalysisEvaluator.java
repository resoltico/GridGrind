package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.InspectionQuery;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.WorkbookAnalysisResult;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookLocation;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Evaluates analysis-summary and finding-presence assertion families. */
final class AssertionAnalysisEvaluator {
  private final AssertionObservationExecutor observations;

  AssertionAnalysisEvaluator(AssertionObservationExecutor observations) {
    this.observations = Objects.requireNonNull(observations, "observations must not be null");
  }

  AssertionEvaluation evaluateAnalysisMaxSeverity(
      StepId stepId,
      Selector target,
      InspectionQuery query,
      AnalysisSeverity maximumSeverity,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation) {
    WorkbookAnalysisResult result =
        (WorkbookAnalysisResult)
            observations.executeObservation(stepId, target, query, workbook, workbookLocation);
    Optional<AnalysisSeverity> observedSeverity = highestSeverity(result);
    return severityRank(observedSeverity) <= severityRank(Optional.of(maximumSeverity))
        ? AssertionEvaluation.pass(List.of(result))
        : AssertionEvaluation.fail(
            List.of(result),
            "EXPECT_ANALYSIS_MAX_SEVERITY observed highest severity "
                + observedSeverity.map(Enum::name).orElse("NONE")
                + " which exceeds "
                + maximumSeverity);
  }

  AssertionEvaluation evaluateFindingPresence(
      StepId stepId,
      Selector target,
      InspectionQuery query,
      AnalysisFindingCode code,
      Optional<AnalysisSeverity> severity,
      Optional<String> messageContains,
      boolean shouldExist,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation) {
    WorkbookAnalysisResult result =
        (WorkbookAnalysisResult)
            observations.executeObservation(stepId, target, query, workbook, workbookLocation);
    boolean found =
        analysisFindings(result).stream()
            .anyMatch(finding -> matchesFinding(finding, code, severity, messageContains));
    if (found == shouldExist) {
      return AssertionEvaluation.pass(List.of(result));
    }
    String verb = shouldExist ? "missing" : "unexpectedly present";
    return AssertionEvaluation.fail(
        List.of(result), query.queryType() + " " + verb + " finding " + code);
  }

  static Optional<AnalysisSeverity> highestSeverity(WorkbookAnalysisResult result) {
    AnalysisSummaryReport summary = analysisSummary(result);
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
    return severity.map(AssertionAnalysisEvaluator::severityRank).orElse(-1);
  }

  private static int severityRank(AnalysisSeverity severity) {
    return switch (Objects.requireNonNull(severity, "severity must not be null")) {
      case INFO -> 0;
      case WARNING -> 1;
      case ERROR -> 2;
    };
  }

  static AnalysisSummaryReport analysisSummary(WorkbookAnalysisResult result) {
    return switch (result) {
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.FormulaHealthResult
              formulaHealthResult ->
          formulaHealthResult.analysis().summary();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.DataValidationHealthResult
              dataValidationHealthResult ->
          dataValidationHealthResult.analysis().summary();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.ConditionalFormattingHealthResult
              conditionalFormattingHealthResult ->
          conditionalFormattingHealthResult.analysis().summary();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.AutofilterHealthResult
              autofilterHealthResult ->
          autofilterHealthResult.analysis().summary();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.TableHealthResult
              tableHealthResult ->
          tableHealthResult.analysis().summary();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.PivotTableHealthResult
              pivotTableHealthResult ->
          pivotTableHealthResult.analysis().summary();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.HyperlinkHealthResult
              hyperlinkHealthResult ->
          hyperlinkHealthResult.analysis().summary();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.NamedRangeHealthResult
              namedRangeHealthResult ->
          namedRangeHealthResult.analysis().summary();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.WorkbookFindingsResult
              workbookFindingsResult ->
          workbookFindingsResult.analysis().summary();
    };
  }

  static List<AnalysisFindingReport> analysisFindings(WorkbookAnalysisResult result) {
    return switch (result) {
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.FormulaHealthResult
              formulaHealthResult ->
          formulaHealthResult.analysis().findings();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.DataValidationHealthResult
              dataValidationHealthResult ->
          dataValidationHealthResult.analysis().findings();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.ConditionalFormattingHealthResult
              conditionalFormattingHealthResult ->
          conditionalFormattingHealthResult.analysis().findings();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.AutofilterHealthResult
              autofilterHealthResult ->
          autofilterHealthResult.analysis().findings();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.TableHealthResult
              tableHealthResult ->
          tableHealthResult.analysis().findings();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.PivotTableHealthResult
              pivotTableHealthResult ->
          pivotTableHealthResult.analysis().findings();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.HyperlinkHealthResult
              hyperlinkHealthResult ->
          hyperlinkHealthResult.analysis().findings();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.NamedRangeHealthResult
              namedRangeHealthResult ->
          namedRangeHealthResult.analysis().findings();
      case dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.WorkbookFindingsResult
              workbookFindingsResult ->
          workbookFindingsResult.analysis().findings();
    };
  }

  static boolean matchesFinding(
      AnalysisFindingReport finding,
      AnalysisFindingCode code,
      Optional<AnalysisSeverity> severity,
      Optional<String> messageContains) {
    if (finding.code() != code) {
      return false;
    }
    if (severity.isPresent() && finding.severity() != severity.orElseThrow()) {
      return false;
    }
    return messageContains.isEmpty() || finding.message().contains(messageContains.orElseThrow());
  }
}
