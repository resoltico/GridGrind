package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.ProtocolCellAddressValidation;
import dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidation;
import dev.erst.gridgrind.excel.foundation.ExcelColumnWidthViolation;
import dev.erst.gridgrind.excel.foundation.ExcelRowHeightViolation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNameProblem;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNameRule;
import dev.erst.gridgrind.excel.foundation.ExcelZoomViolation;

/** Converts upstream typed validation failures into request-problem descriptors. */
public final class FieldValidationProblemMappers {
  private FieldValidationProblemMappers() {}

  /** Maps one typed sheet-name failure onto the public request-problem surface. */
  public static FieldValidationProblem sheetName(String fieldName, ExcelSheetNameProblem problem) {
    if (problem.rule() == ExcelSheetNameRule.BLANK) {
      return FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_BLANK);
    }
    if (problem.rule() == ExcelSheetNameRule.TOO_LONG) {
      return FieldValidationProblem.atField(
          fieldName, FieldValidationNamingRule.SHEET_NAME_TOO_LONG, problem.operand(0));
    }
    if (problem.rule() == ExcelSheetNameRule.BOUNDARY_QUOTE) {
      return FieldValidationProblem.atField(
          fieldName, FieldValidationNamingRule.SHEET_NAME_BOUNDARY_QUOTE, problem.operand(0));
    }
    return FieldValidationProblem.atField(
        fieldName,
        FieldValidationNamingRule.SHEET_NAME_INVALID_CHARACTER,
        problem.operand(0),
        problem.operand(1),
        problem.operand(2));
  }

  /** Maps one typed defined-name failure onto the public request-problem surface. */
  public static FieldValidationProblem definedName(
      String fieldName, ProtocolDefinedNameValidation.Violation violation) {
    if (violation == ProtocolDefinedNameValidation.Violation.BLANK) {
      return FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_BLANK);
    }
    if (violation == ProtocolDefinedNameValidation.Violation.TOO_LONG) {
      return FieldValidationProblem.atField(
          fieldName, FieldValidationNamingRule.DEFINED_NAME_TOO_LONG);
    }
    if (violation == ProtocolDefinedNameValidation.Violation.SYNTAX) {
      return FieldValidationProblem.atField(
          fieldName, FieldValidationNamingRule.DEFINED_NAME_SYNTAX);
    }
    if (violation == ProtocolDefinedNameValidation.Violation.RESERVED_PREFIX) {
      return FieldValidationProblem.atField(
          fieldName, FieldValidationNamingRule.DEFINED_NAME_RESERVED_PREFIX);
    }
    if (violation == ProtocolDefinedNameValidation.Violation.A1_COLLISION) {
      return FieldValidationProblem.atField(
          fieldName, FieldValidationNamingRule.DEFINED_NAME_A1_COLLISION);
    }
    return FieldValidationProblem.atField(
        fieldName, FieldValidationNamingRule.DEFINED_NAME_R1C1_COLLISION);
  }

  /** Maps one typed single-cell address failure onto the public request-problem surface. */
  public static FieldValidationProblem address(
      String fieldName, ProtocolCellAddressValidation.Violation violation) {
    if (violation == ProtocolCellAddressValidation.Violation.BLANK) {
      return FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_BLANK);
    }
    if (violation == ProtocolCellAddressValidation.Violation.SYNTAX) {
      return FieldValidationProblem.atField(fieldName, FieldValidationAddressRule.ADDRESS_SYNTAX);
    }
    return FieldValidationProblem.atField(fieldName, FieldValidationAddressRule.ADDRESS_BOUNDS);
  }

  /** Maps the pivot-table blank-name failure onto the public request-problem surface. */
  public static FieldValidationProblem pivotTableName(String fieldName) {
    return FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_BLANK);
  }

  /** Maps one typed column-width failure onto the public request-problem surface. */
  public static FieldValidationProblem columnWidth(
      String fieldName, double widthCharacters, ExcelColumnWidthViolation violation) {
    if (violation == ExcelColumnWidthViolation.NON_FINITE) {
      return FieldValidationProblem.atField(
          fieldName, FieldValidationLayoutRule.FIELD_MUST_BE_FINITE);
    }
    if (violation == ExcelColumnWidthViolation.NON_POSITIVE) {
      return FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.GREATER_THAN_ZERO);
    }
    if (violation == ExcelColumnWidthViolation.TOO_LARGE) {
      return FieldValidationProblem.atField(
          fieldName,
          FieldValidationLayoutRule.COLUMN_WIDTH_TOO_LARGE,
          Double.toString(ExcelSheetLayoutLimits.MAX_COLUMN_WIDTH_CHARACTERS),
          Double.toString(widthCharacters));
    }
    return FieldValidationProblem.atField(
        fieldName,
        FieldValidationLayoutRule.COLUMN_WIDTH_NOT_VISIBLE,
        Double.toString(ExcelSheetLayoutLimits.MAX_COLUMN_WIDTH_CHARACTERS),
        Double.toString(widthCharacters));
  }

  /** Maps one typed row-height failure onto the public request-problem surface. */
  public static FieldValidationProblem rowHeight(
      String fieldName, double heightPoints, ExcelRowHeightViolation violation) {
    if (violation == ExcelRowHeightViolation.NON_FINITE) {
      return FieldValidationProblem.atField(
          fieldName, FieldValidationLayoutRule.FIELD_MUST_BE_FINITE);
    }
    if (violation == ExcelRowHeightViolation.NON_POSITIVE) {
      return FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.GREATER_THAN_ZERO);
    }
    if (violation == ExcelRowHeightViolation.TOO_LARGE) {
      return FieldValidationProblem.atField(
          fieldName,
          FieldValidationLayoutRule.ROW_HEIGHT_TOO_LARGE,
          Double.toString(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS),
          Double.toString(heightPoints));
    }
    return FieldValidationProblem.atField(
        fieldName,
        FieldValidationLayoutRule.ROW_HEIGHT_NOT_VISIBLE,
        Double.toString(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS),
        Double.toString(heightPoints));
  }

  /** Maps one typed zoom-range failure onto the public request-problem surface. */
  public static FieldValidationProblem zoomPercent(
      String fieldName, int zoomPercent, ExcelZoomViolation violation) {
    if (violation != ExcelZoomViolation.OUT_OF_RANGE) {
      throw new IllegalArgumentException("Unhandled zoom violation: " + violation);
    }
    return FieldValidationProblem.atField(
        fieldName,
        FieldValidationLayoutRule.ZOOM_PERCENT_RANGE,
        Integer.toString(ExcelSheetLayoutLimits.MIN_ZOOM_PERCENT),
        Integer.toString(ExcelSheetLayoutLimits.MAX_ZOOM_PERCENT),
        Integer.toString(zoomPercent));
  }
}
