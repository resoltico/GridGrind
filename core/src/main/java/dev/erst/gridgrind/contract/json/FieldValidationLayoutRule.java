package dev.erst.gridgrind.contract.json;

/** Layout and sizing rules tied to Excel workbook presentation limits. */
public enum FieldValidationLayoutRule implements FieldValidationRule {
  FIELD_MUST_BE_FINITE,
  COLUMN_WIDTH_TOO_LARGE,
  COLUMN_WIDTH_NOT_VISIBLE,
  ROW_HEIGHT_TOO_LARGE,
  ROW_HEIGHT_NOT_VISIBLE,
  ZOOM_PERCENT_RANGE;

  @Override
  public String message(FieldValidationProblem problem) {
    return switch (this) {
      case FIELD_MUST_BE_FINITE -> problem.fieldName() + " must be finite";
      case COLUMN_WIDTH_TOO_LARGE ->
          problem.fieldName()
              + " must not exceed "
              + problem.operand(0)
              + " (Excel column width limit): got "
              + problem.operand(1);
      case COLUMN_WIDTH_NOT_VISIBLE ->
          problem.fieldName()
              + " is too small to produce a visible Excel column width: got "
              + problem.operand(1);
      case ROW_HEIGHT_TOO_LARGE ->
          problem.fieldName()
              + " must not exceed "
              + problem.operand(0)
              + " (Excel row height limit): got "
              + problem.operand(1);
      case ROW_HEIGHT_NOT_VISIBLE ->
          problem.fieldName()
              + " is too small to produce a visible Excel row height: "
              + problem.operand(1);
      case ZOOM_PERCENT_RANGE ->
          problem.fieldName()
              + " must be between "
              + problem.operand(0)
              + " and "
              + problem.operand(1)
              + " inclusive: "
              + problem.operand(2);
    };
  }

  @Override
  public String resolution(FieldValidationProblem problem) {
    return switch (this) {
      case FIELD_MUST_BE_FINITE ->
          "Provide a finite numeric value for field '" + problem.fieldName() + "'.";
      case COLUMN_WIDTH_TOO_LARGE, COLUMN_WIDTH_NOT_VISIBLE ->
          "Provide a visible Excel column width greater than 0 and no more than "
              + problem.operand(0)
              + " characters for field '"
              + problem.fieldName()
              + "'.";
      case ROW_HEIGHT_TOO_LARGE, ROW_HEIGHT_NOT_VISIBLE ->
          "Provide a visible Excel row height greater than 0 and no more than "
              + problem.operand(0)
              + " points for field '"
              + problem.fieldName()
              + "'.";
      case ZOOM_PERCENT_RANGE ->
          "Provide a zoom percentage between "
              + problem.operand(0)
              + " and "
              + problem.operand(1)
              + " inclusive for field '"
              + problem.fieldName()
              + "'.";
    };
  }
}
