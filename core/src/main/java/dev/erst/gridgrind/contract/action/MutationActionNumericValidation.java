package dev.erst.gridgrind.contract.action;

import dev.erst.gridgrind.contract.json.FieldValidationBasicRule;
import dev.erst.gridgrind.contract.json.FieldValidationBoundRule;
import dev.erst.gridgrind.contract.json.FieldValidationProblem;
import dev.erst.gridgrind.contract.json.FieldValidationProblemMappers;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;

/** Numeric and span validation for mutation actions. */
final class MutationActionNumericValidation {
  private MutationActionNumericValidation() {}

  static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw MutationActionNameValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_NEGATIVE));
    }
  }

  static void requirePositive(int value, String fieldName) {
    if (value <= 0) {
      throw MutationActionNameValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.GREATER_THAN_ZERO));
    }
  }

  static void requireNonZero(int value, String fieldName) {
    if (value == 0) {
      throw MutationActionNameValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_ZERO));
    }
  }

  static void requireRowIndex(int value, String fieldName) {
    requireNonNegative(value, fieldName);
    if (value > ExcelRowSpan.MAX_ROW_INDEX) {
      throw MutationActionNameValidation.invalidField(
          FieldValidationProblem.atField(
              fieldName,
              FieldValidationBoundRule.ROW_INDEX_BOUNDS,
              Integer.toString(ExcelRowSpan.MAX_ROW_INDEX)));
    }
  }

  static void requireColumnIndex(int value, String fieldName) {
    requireNonNegative(value, fieldName);
    if (value > ExcelColumnSpan.MAX_COLUMN_INDEX) {
      throw MutationActionNameValidation.invalidField(
          FieldValidationProblem.atField(
              fieldName,
              FieldValidationBoundRule.COLUMN_INDEX_BOUNDS,
              Integer.toString(ExcelColumnSpan.MAX_COLUMN_INDEX)));
    }
  }

  static void requireOrderedSpan(
      int firstValue, int lastValue, String firstFieldName, String lastFieldName) {
    if (lastValue < firstValue) {
      throw MutationActionNameValidation.invalidField(
          FieldValidationProblem.atField(
              lastFieldName, FieldValidationBoundRule.ORDERED_SPAN, firstFieldName));
    }
  }

  static void requireColumnWidthCharacters(double widthCharacters) {
    ExcelSheetLayoutLimits.columnWidthViolation(widthCharacters)
        .ifPresent(
            violation -> {
              throw MutationActionNameValidation.invalidField(
                  FieldValidationProblemMappers.columnWidth(
                      "widthCharacters", widthCharacters, violation));
            });
  }

  static void requireRowHeightPoints(double heightPoints) {
    ExcelSheetLayoutLimits.rowHeightViolation(heightPoints)
        .ifPresent(
            violation -> {
              throw MutationActionNameValidation.invalidField(
                  FieldValidationProblemMappers.rowHeight("heightPoints", heightPoints, violation));
            });
  }

  static void requireZoomPercent(int zoomPercent) {
    ExcelSheetLayoutLimits.zoomViolation(zoomPercent)
        .ifPresent(
            violation -> {
              throw MutationActionNameValidation.invalidField(
                  FieldValidationProblemMappers.zoomPercent("zoomPercent", zoomPercent, violation));
            });
  }
}
