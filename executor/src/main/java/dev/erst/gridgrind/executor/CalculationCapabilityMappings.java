package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.FormulaCapabilityKind;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.excel.ExcelFormulaCapabilityAssessment;
import dev.erst.gridgrind.excel.ExcelFormulaCapabilityKind;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Canonical executor mappings for engine-side formula capability facts. */
final class CalculationCapabilityMappings {
  private CalculationCapabilityMappings() {}

  static CalculationReport.Summary summaryFor(List<ExcelFormulaCapabilityAssessment> assessments) {
    int evaluableNowCount =
        Math.toIntExact(
            assessments.stream()
                .filter(
                    assessment ->
                        assessment.capability() == ExcelFormulaCapabilityKind.EVALUABLE_NOW)
                .count());
    int unevaluableNowCount =
        Math.toIntExact(
            assessments.stream()
                .filter(
                    assessment ->
                        assessment.capability() == ExcelFormulaCapabilityKind.UNEVALUABLE_NOW)
                .count());
    int unparseableByPoiCount =
        Math.toIntExact(
            assessments.stream()
                .filter(
                    assessment ->
                        assessment.capability() == ExcelFormulaCapabilityKind.UNPARSEABLE_BY_POI)
                .count());
    return new CalculationReport.Summary(
        evaluableNowCount, unevaluableNowCount, unparseableByPoiCount);
  }

  static FormulaCapabilityKind capabilityKindFor(ExcelFormulaCapabilityKind capability) {
    return switch (capability) {
      case EVALUABLE_NOW -> FormulaCapabilityKind.EVALUABLE_NOW;
      case UNEVALUABLE_NOW -> FormulaCapabilityKind.UNEVALUABLE_NOW;
      case UNPARSEABLE_BY_POI -> FormulaCapabilityKind.UNPARSEABLE_BY_POI;
    };
  }

  static Optional<GridGrindProblemCode> problemCodeFor(
      ExcelFormulaCapabilityAssessment assessment) {
    Objects.requireNonNull(assessment, "assessment must not be null");
    if (assessment.issue() == null) {
      return Optional.empty();
    }
    return Optional.of(
        switch (assessment.issue()) {
          case INVALID_FORMULA -> GridGrindProblemCode.INVALID_FORMULA;
          case MISSING_EXTERNAL_WORKBOOK -> GridGrindProblemCode.MISSING_EXTERNAL_WORKBOOK;
          case UNREGISTERED_USER_DEFINED_FUNCTION ->
              GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION;
          case UNSUPPORTED_FORMULA -> GridGrindProblemCode.UNSUPPORTED_FORMULA;
        });
  }

  static int severityRank(ExcelFormulaCapabilityAssessment assessment) {
    Objects.requireNonNull(assessment, "assessment must not be null");
    var issue = Objects.requireNonNull(assessment.issue(), "assessment.issue must not be null");
    return switch (issue) {
      case INVALID_FORMULA -> 0;
      case MISSING_EXTERNAL_WORKBOOK -> 1;
      case UNREGISTERED_USER_DEFINED_FUNCTION -> 2;
      case UNSUPPORTED_FORMULA -> 3;
    };
  }
}
