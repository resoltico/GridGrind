package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.model.formula.FormulaCapabilityKind;
import java.util.List;
import java.util.Objects;

/** Aggregates canonical formula capabilities and orders blocking calculation failures. */
final class CalculationCapabilities {
  private CalculationCapabilities() {}

  static CalculationReport.Summary summaryFor(
      List<CalculationReport.FormulaCapability> assessments) {
    int evaluableNowCount =
        Math.toIntExact(
            assessments.stream()
                .filter(
                    assessment -> assessment.capability() == FormulaCapabilityKind.EVALUABLE_NOW)
                .count());
    int unevaluableNowCount =
        Math.toIntExact(
            assessments.stream()
                .filter(
                    assessment -> assessment.capability() == FormulaCapabilityKind.UNEVALUABLE_NOW)
                .count());
    int unparseableByPoiCount =
        Math.toIntExact(
            assessments.stream()
                .filter(
                    assessment ->
                        assessment.capability() == FormulaCapabilityKind.UNPARSEABLE_BY_POI)
                .count());
    return new CalculationReport.Summary(
        evaluableNowCount, unevaluableNowCount, unparseableByPoiCount);
  }

  static int severityRank(CalculationReport.FormulaCapability assessment) {
    Objects.requireNonNull(assessment, "assessment must not be null");
    var issue = assessment.problemCode().orElseThrow();
    return switch (issue) {
      case INVALID_FORMULA -> 0;
      case CIRCULAR_FORMULA_REFERENCE -> 1;
      case MISSING_EXTERNAL_WORKBOOK -> 2;
      case UNREGISTERED_USER_DEFINED_FUNCTION -> 3;
      case UNSUPPORTED_FORMULA -> 4;
      default -> throw new IllegalArgumentException("Not a formula capability problem: " + issue);
    };
  }
}
