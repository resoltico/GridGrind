package dev.erst.gridgrind.contract.json;

/** Excel naming rules whose wording depends on authored names and name syntax. */
public enum FieldValidationNamingRule implements FieldValidationRule {
  SHEET_NAME_TOO_LONG,
  SHEET_NAME_BOUNDARY_QUOTE,
  SHEET_NAME_INVALID_CHARACTER,
  DEFINED_NAME_TOO_LONG,
  DEFINED_NAME_SYNTAX,
  DEFINED_NAME_RESERVED_PREFIX,
  DEFINED_NAME_A1_COLLISION,
  DEFINED_NAME_R1C1_COLLISION;

  @Override
  public String message(FieldValidationProblem problem) {
    return switch (this) {
      case SHEET_NAME_TOO_LONG ->
          problem.fieldName() + " must not exceed 31 characters: " + problem.operand(0);
      case SHEET_NAME_BOUNDARY_QUOTE ->
          problem.fieldName() + " must not begin or end with a single quote: " + problem.operand(0);
      case SHEET_NAME_INVALID_CHARACTER ->
          problem.fieldName()
              + " contains invalid Excel character "
              + problem.operand(0)
              + " at position "
              + problem.operand(1)
              + ": "
              + problem.operand(2);
      case DEFINED_NAME_TOO_LONG ->
          problem.fieldName() + " must not exceed 255 Unicode code points";
      case DEFINED_NAME_SYNTAX ->
          problem.fieldName()
              + " must start with a letter, underscore, or backslash and contain only Unicode"
              + " letters, Unicode numbers, underscore, period, or backslash";
      case DEFINED_NAME_RESERVED_PREFIX ->
          problem.fieldName() + " must not use the reserved _xlnm. prefix";
      case DEFINED_NAME_A1_COLLISION ->
          problem.fieldName() + " must not collide with A1-style cell reference syntax";
      case DEFINED_NAME_R1C1_COLLISION ->
          problem.fieldName() + " must not collide with R1C1-style cell reference syntax";
    };
  }

  @Override
  public String resolution(FieldValidationProblem problem) {
    return switch (this) {
      case SHEET_NAME_TOO_LONG, SHEET_NAME_BOUNDARY_QUOTE, SHEET_NAME_INVALID_CHARACTER ->
          "Provide a valid Excel sheet name for field '" + problem.fieldName() + "'.";
      case DEFINED_NAME_TOO_LONG,
          DEFINED_NAME_SYNTAX,
          DEFINED_NAME_RESERVED_PREFIX,
          DEFINED_NAME_A1_COLLISION,
          DEFINED_NAME_R1C1_COLLISION ->
          "Provide a valid Excel defined name for field '" + problem.fieldName() + "'.";
    };
  }
}
