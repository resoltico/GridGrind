package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport;
import java.nio.charset.StandardCharsets;

/** Owns request-specific analysis, early structural rejection, and typed DTO decoding. */
public final class GridGrindRequestDecoder {
  private GridGrindRequestDecoder() {}

  /** Binds one request document through tolerant intake. */
  public static WorkbookPlan read(byte[] bytes) {
    return completePlan(analyze(bytes));
  }

  /** Binds one request document through tolerant intake. */
  public static WorkbookPlan read(String json) {
    return completePlan(analyze(json.getBytes(StandardCharsets.UTF_8)));
  }

  private static WorkbookPlan completePlan(RequestIntake analysis) {
    analysis.structuralProblems().stream()
        .findFirst()
        .ifPresent(
            problem -> {
              throw problem.exception();
            });
    return analysis.requireCompletePlan();
  }

  /** Collects independently bound request fragments and ordered findings. */
  public static RequestIntake analyze(byte[] bytes) {
    GridGrindJsonMapperSupport.requireSupportedRequestLength(bytes.length);
    return RequestStructuralAnalyzer.analyze(bytes);
  }
}
