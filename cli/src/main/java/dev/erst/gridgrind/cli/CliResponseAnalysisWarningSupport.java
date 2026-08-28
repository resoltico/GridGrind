package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import java.util.ArrayList;
import java.util.List;

/** Appends non-fatal request-transport warnings without changing a logical workbook result. */
final class CliResponseAnalysisWarningSupport {
  private CliResponseAnalysisWarningSupport() {}

  static WorkbookResult append(WorkbookResult response, RequestAnalysis analysis) {
    return append(response, analysis.warnings());
  }

  static WorkbookResult append(WorkbookResult response, List<RequestWarning> analysisWarnings) {
    if (analysisWarnings.isEmpty()) {
      return response;
    }
    List<RequestWarning> warnings = new ArrayList<>();
    return switch (response) {
      case WorkbookResult.Success success -> {
        warnings.addAll(success.warnings());
        warnings.addAll(analysisWarnings);
        yield new WorkbookResult.Success(
            success.protocolVersion(),
            success.planId(),
            success.journal(),
            success.calculation(),
            success.persistence(),
            warnings,
            success.assertions(),
            success.inspections());
      }
      case WorkbookResult.Failure failure -> {
        warnings.addAll(failure.warnings());
        warnings.addAll(analysisWarnings);
        yield new WorkbookResult.Failure(
            failure.protocolVersion(),
            failure.planId(),
            failure.journal(),
            failure.calculation(),
            failure.persistence(),
            warnings,
            failure.assertions(),
            failure.inspections(),
            failure.problem());
      }
    };
  }
}
