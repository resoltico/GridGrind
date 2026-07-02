package dev.erst.gridgrind.contract.action;

import dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidation;
import dev.erst.gridgrind.contract.json.FieldValidationBasicRule;
import dev.erst.gridgrind.contract.json.FieldValidationProblem;
import dev.erst.gridgrind.contract.json.FieldValidationProblemMappers;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.excel.foundation.ExcelPivotTableNaming;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.Objects;
import java.util.Optional;

/** Text and naming validation for mutation actions. */
final class MutationActionNameValidation {
  private MutationActionNameValidation() {}

  static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_BLANK));
    }
  }

  static void requireSheetName(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    ExcelSheetNames.violation(value)
        .ifPresent(
            violation -> {
              throw invalidField(FieldValidationProblemMappers.sheetName(fieldName, violation));
            });
  }

  static void requireNamedRangeName(String name) {
    Objects.requireNonNull(name, "name must not be null");
    ProtocolDefinedNameValidation.violation(name)
        .ifPresent(
            violation -> {
              throw invalidField(FieldValidationProblemMappers.definedName("name", violation));
            });
  }

  static void requirePivotTableName(String name) {
    Objects.requireNonNull(name, "name must not be null");
    ExcelPivotTableNaming.violation(name)
        .ifPresent(
            violation -> {
              throw invalidField(FieldValidationProblemMappers.pivotTableName("name"));
            });
  }

  static InvalidRequestException invalidField(FieldValidationProblem problem) {
    return new InvalidRequestException(
        problem, Optional.empty(), Optional.empty(), Optional.empty(), null);
  }
}
