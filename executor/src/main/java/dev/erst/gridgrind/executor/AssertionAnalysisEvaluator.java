package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.GridGrindAnalysisReports;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
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
      String stepId,
      Selector target,
      InspectionQuery query,
      AnalysisSeverity maximumSeverity,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.Analysis result =
        (InspectionResult.Analysis)
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
      String stepId,
      Selector target,
      InspectionQuery query,
      AnalysisFindingCode code,
      AnalysisSeverity severity,
      String messageContains,
      boolean shouldExist,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation) {
    InspectionResult.Analysis result =
        (InspectionResult.Analysis)
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

  static Optional<AnalysisSeverity> highestSeverity(InspectionResult.Analysis result) {
    GridGrindAnalysisReports.AnalysisSummaryReport summary = analysisSummary(result);
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

  static GridGrindAnalysisReports.AnalysisSummaryReport analysisSummary(
      InspectionResult.Analysis result) {
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

  static List<GridGrindAnalysisReports.AnalysisFindingReport> analysisFindings(
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
      GridGrindAnalysisReports.AnalysisFindingReport finding,
      AnalysisFindingCode code,
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
}
