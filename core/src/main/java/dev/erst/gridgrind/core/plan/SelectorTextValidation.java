package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.problem.FieldValidationBasicRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.FieldValidationProblemMappers;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotTableNaming;
import java.util.Objects;
import java.util.Optional;

/** Text and identifier validation for selector families. */
final class SelectorTextValidation {
  private SelectorTextValidation() {}

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_BLANK));
    }
    return value;
  }

  static String requirePivotTableName(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    ExcelPivotTableNaming.violation(value)
        .ifPresent(violation -> {
          throw invalidField(FieldValidationProblemMappers.pivotTableName(fieldName));
        });
    return value;
  }

  static dev.erst.gridgrind.core.vocabulary.RangeRef requireOrderedRange(
      dev.erst.gridgrind.core.vocabulary.RangeRef value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    return value;
  }

  static InvalidRequestException invalidField(FieldValidationProblem problem) {
    return new InvalidRequestException(
        problem, Optional.empty(), Optional.empty(), Optional.empty(), null);
  }
}
