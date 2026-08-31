package dev.erst.gridgrind.contract.selector;

import dev.erst.gridgrind.contract.dto.ProtocolCellAddressValidation;
import dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidation;
import dev.erst.gridgrind.contract.json.FieldValidationAddressRule;
import dev.erst.gridgrind.contract.json.FieldValidationBasicRule;
import dev.erst.gridgrind.contract.json.FieldValidationProblem;
import dev.erst.gridgrind.contract.json.FieldValidationProblemMappers;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.excel.foundation.ExcelPivotTableNaming;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
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

  static String requireSheetName(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    ExcelSheetNames.violation(value)
        .ifPresent(
            violation -> {
              throw invalidField(FieldValidationProblemMappers.sheetName(fieldName, violation));
            });
    return value;
  }

  static String requireDefinedName(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    ProtocolDefinedNameValidation.violation(value)
        .ifPresent(
            violation -> {
              throw invalidField(FieldValidationProblemMappers.definedName(fieldName, violation));
            });
    return value;
  }

  static String requirePivotTableName(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    ExcelPivotTableNaming.violation(value)
        .ifPresent(
            violation -> {
              throw invalidField(FieldValidationProblemMappers.pivotTableName(fieldName));
            });
    return value;
  }

  static String requireAddress(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    ProtocolCellAddressValidation.violation(value)
        .ifPresent(
            violation -> {
              throw invalidField(FieldValidationProblemMappers.address(fieldName, violation));
            });
    return value;
  }

  static String requireRange(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_BLANK));
    }
    String[] parts = value.split(":", -1);
    if (parts.length > 2) {
      throw invalidField(
          FieldValidationProblem.atField(
              fieldName, FieldValidationAddressRule.RANGE_RECTANGULAR_SYNTAX));
    }
    requireAddress(parts[0], fieldName);
    if (parts.length == 2) {
      requireAddress(parts[1], fieldName);
      if (SelectorAddressSupport.rowIndex(parts[1]) < SelectorAddressSupport.rowIndex(parts[0])
          || SelectorAddressSupport.columnIndex(parts[1])
              < SelectorAddressSupport.columnIndex(parts[0])) {
        throw invalidField(
            FieldValidationProblem.atField(fieldName, FieldValidationAddressRule.RANGE_ORDER));
      }
    }
    return value;
  }

  static InvalidRequestException invalidField(FieldValidationProblem problem) {
    return new InvalidRequestException(
        problem, Optional.empty(), Optional.empty(), Optional.empty(), null);
  }
}
