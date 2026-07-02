package dev.erst.gridgrind.contract.json;

/** Scalar and collection rules whose wording does not depend on external limits. */
public enum FieldValidationBasicRule implements FieldValidationRule {
  NON_BLANK,
  NON_EMPTY,
  ROWS_NON_EMPTY,
  NON_NEGATIVE,
  GREATER_THAN_ZERO,
  NON_ZERO,
  DUPLICATES,
  EMPTY_ROWS;

  @Override
  public String message(FieldValidationProblem problem) {
    return switch (this) {
      case NON_BLANK -> problem.fieldName() + " must not be blank";
      case NON_EMPTY -> problem.fieldName() + " must not be empty";
      case ROWS_NON_EMPTY -> problem.fieldName() + " must not be empty";
      case NON_NEGATIVE -> problem.fieldName() + " must not be negative";
      case GREATER_THAN_ZERO -> problem.fieldName() + " must be greater than 0";
      case NON_ZERO -> problem.fieldName() + " must not be 0";
      case DUPLICATES -> problem.fieldName() + " must not contain duplicates";
      case EMPTY_ROWS -> problem.fieldName() + " must not contain empty rows";
    };
  }

  @Override
  public String resolution(FieldValidationProblem problem) {
    return switch (this) {
      case NON_BLANK -> "Provide a non-blank value for field '" + problem.fieldName() + "'.";
      case NON_EMPTY -> "Provide at least one value in field '" + problem.fieldName() + "'.";
      case ROWS_NON_EMPTY ->
          "Provide at least one non-empty row in field '" + problem.fieldName() + "'.";
      case NON_NEGATIVE ->
          "Provide a non-negative integer for field '" + problem.fieldName() + "'.";
      case GREATER_THAN_ZERO ->
          "Provide an integer greater than 0 for field '" + problem.fieldName() + "'.";
      case NON_ZERO -> "Provide a non-zero integer for field '" + problem.fieldName() + "'.";
      case DUPLICATES -> "Remove duplicate values from field '" + problem.fieldName() + "'.";
      case EMPTY_ROWS ->
          "Ensure every row in field '"
              + problem.fieldName()
              + "' contains at least one cell value.";
    };
  }
}
