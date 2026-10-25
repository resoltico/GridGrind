package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.nio.charset.StandardCharsets;

/** Owns request-specific analysis, early structural rejection, and typed DTO decoding. */
final class GridGrindRequestDecoder {
  private GridGrindRequestDecoder() {}

  static WorkbookPlan read(byte[] bytes) {
    return completePlan(analyze(bytes));
  }

  static WorkbookPlan read(String json) {
    return completePlan(analyze(json.getBytes(StandardCharsets.UTF_8)));
  }

  private static WorkbookPlan completePlan(RequestAnalysis analysis) {
    analysis.structuralProblems().stream()
        .findFirst()
        .ifPresent(
            problem -> {
              throw problem.exception();
            });
    return analysis.requireCompletePlan();
  }

  static RequestAnalysis analyze(byte[] bytes) {
    GridGrindJsonMapperSupport.requireSupportedRequestLength(bytes.length);
    return RequestStructuralAnalyzer.analyze(bytes);
  }
}
